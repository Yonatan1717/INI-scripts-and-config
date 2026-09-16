import os
import sys
import json
import pandas as pd
from openpyxl import load_workbook
import ipaddress

DEFAULT_IPSEC_PSK = "DMVPN-KEY"
SSH_DOMAIN = "lab.local"
FLOW_POLICY_SHEET = "FLOW_POLICY"
FLOW_POLICY_COLUMNS = ["source", "destination", "protocol", "destination_port", "action"]


def read_sheet(filename, sheet):
    df = pd.read_excel(
        filename,
        sheet_name=sheet,
        header=None
    )

    md_start = 0
    md_end = df.iloc[md_start:].isna().all(axis=1).idxmax()
    # Metadata width is intentionally dynamic so new product options (for
    # example dns_servers) can be added in Excel without changing this parser.
    md = df.iloc[md_start:md_end].dropna(axis=1, how="all")
    md.columns = md.iloc[0]
    md = md[1:].reset_index(drop=True)

    ip_data_start = md_end + 1
    blank = df.iloc[ip_data_start:].isna().all(axis=1)
    ip_data_end = blank.idxmax()
    ip_data = df.iloc[ip_data_start:ip_data_end, 0:11]
    ip_data.columns = ip_data.iloc[0]
    ip_data = ip_data[1:].reset_index(drop=True)


    vrf_data_start = ip_data_end + 1
    blank = df.iloc[vrf_data_start:].isna().all(axis=1)
    vrf_data_end = blank.idxmax()
    vrf_data = df.iloc[vrf_data_start:vrf_data_end, 0:4]
    vrf_data.columns = vrf_data.iloc[0]
    vrf_data = vrf_data[1:].reset_index(drop=True)
    
    tunnel_data_start = vrf_data_end + 1
    blank = df.iloc[tunnel_data_start:].isna().all(axis=1)
    tunnel_data_end = blank.idxmax()
    tunnel_data = df.iloc[tunnel_data_start:tunnel_data_end, 0:10]
    tunnel_data.columns = tunnel_data.iloc[0]
    tunnel_data = tunnel_data[1:].reset_index(drop=True)

    # print(md)
    # print(ip_data)
    # print(vrf_data)
    # print(tunnel_data)
    # exit(0)
    
    return {
        "md": md,
        "ip_data": ip_data,
        "vrf_data": vrf_data,
        "tunnel_data": tunnel_data,
    }

    # Removed redundant return statement


def is_true(value):
    """
    Godta True/TRUE/1/yes/ja/x fra sheets.
    """
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value == 1

    return str(value).strip().lower() in {"true", "1", "yes", "ja", "y", "x"}


def unique_preserve_order(values):
    """Fjern duplikater uten å endre CLI-rekkefølgen."""
    result = []
    for value in values:
        if value not in result:
            result.append(value)
    return result


def ensure_exit_last(commands):
    """Sørg for at Cisco `exit` alltid er siste kommando i en CLI-blokk."""
    commands = [cmd for cmd in commands if cmd != "exit"]
    return unique_preserve_order(commands) + ["exit"]


def getNetId(ip, mask):
    ip = ip.split(".")
    mask = mask.split(".")
    netid = []
    wild_mask = [255, 255, 255, 255]
    wild = []

    for i in range(4):
        ip_b = int(ip[i])
        ip_m = int(mask[i])
        ip_w = wild_mask[i]

        wild.append(str(ip_m ^ ip_w))
        netid.append(str(ip_b & ip_m))

    return ".".join(netid), ".".join(mask), ".".join(wild)



def _normalise_cell_text(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _site_sheet_names(filename):
    """Return only site sheets; FLOW_POLICY is shared workbook metadata."""
    return [
        name
        for name in load_workbook(filename, read_only=True).sheetnames
        if str(name).strip().upper() != FLOW_POLICY_SHEET
    ]


def read_flow_policy(filename):
    """Read the optional shared FLOW_POLICY sheet.

    Removing the sheet (or leaving it empty) keeps backwards compatibility and
    simply disables generated data-plane ACLs.
    """
    try:
        policy = pd.read_excel(filename, sheet_name=FLOW_POLICY_SHEET, dtype=object)
    except ValueError:
        return pd.DataFrame(columns=FLOW_POLICY_COLUMNS)

    if policy.empty:
        return pd.DataFrame(columns=FLOW_POLICY_COLUMNS)

    policy.columns = [str(col).strip().lower() for col in policy.columns]
    missing = [col for col in FLOW_POLICY_COLUMNS if col not in policy.columns]
    if missing:
        raise ValueError(
            f"{FLOW_POLICY_SHEET} mangler kolonne(r): {', '.join(missing)}"
        )

    policy = policy[FLOW_POLICY_COLUMNS].dropna(how="all").reset_index(drop=True)
    for col in FLOW_POLICY_COLUMNS:
        policy[col] = policy[col].map(_normalise_cell_text)
    return policy


def _metadata_value(md, names):
    if md is None or md.empty:
        return None
    row = md.iloc[0]
    for name in names:
        if name in row.index:
            value = row.get(name)
            if not pd.isna(value) and str(value).strip():
                return str(value).strip()
    return None


def _append_unique_vrf_endpoint(context, bucket_name, vrf, endpoint):
    """Append a unique endpoint to one VRF-specific context bucket."""
    bucket = context[bucket_name].setdefault(vrf, [])
    if endpoint not in bucket:
        bucket.append(endpoint)


def build_flow_policy_context(filename, site_sheets):
    """Build the workbook-wide address map used by FLOW_POLICY.

    LAN networks and router loopbacks are deliberately kept separate:
      * A VRF name used as a FLOW_POLICY destination means user/LAN networks.
      * Router loopbacks are infrastructure addresses and are NOT implicitly
        included by a VRF-to-VRF policy rule.
      * Loopbacks are still known to the compiler so INTERNET cannot silently
        become a permit to internal router infrastructure.
    """
    context = {
        "vrf_networks": {},
        "vrf_loopbacks": {},
        "dns": [],
        "tacacs": [],
        "syslog": [],
    }

    for sheet in site_sheets:
        sheet_data = read_sheet(filename, sheet)
        ip_data = sheet_data["ip_data"]
        vrf_data = sheet_data["vrf_data"]
        md = sheet_data["md"]

        for _, row in ip_data.iterrows():
            vrf = _normalise_cell_text(row.get("vrf")).upper()
            if not vrf:
                continue
            try:
                network = ipaddress.ip_network(
                    f"{row.get('nett id')}/{row.get('mask')}", strict=False
                )
            except ValueError as exc:
                raise ValueError(
                    f"{sheet}: ugyldig nett i VRF {vrf}: "
                    f"{row.get('nett id')} {row.get('mask')}"
                ) from exc
            _append_unique_vrf_endpoint(
                context,
                "vrf_networks",
                vrf,
                ("network", str(network.network_address), str(network.hostmask)),
            )

        for _, row in vrf_data.iterrows():
            vrf = _normalise_cell_text(row.get("vrf")).upper()
            laddr = _normalise_cell_text(row.get("laddr"))
            if vrf and laddr:
                try:
                    laddr = str(ipaddress.ip_address(laddr))
                except ValueError as exc:
                    raise ValueError(f"{sheet}: ugyldig loopback-adresse {laddr}") from exc
                _append_unique_vrf_endpoint(
                    context, "vrf_loopbacks", vrf, ("host", laddr)
                )

        dns_value = _metadata_value(md, ["dns_servers"])
        if dns_value:
            for server in dns_value.replace(",", " ").replace(";", " ").split():
                server = str(ipaddress.ip_address(server))
                if server not in context["dns"]:
                    context["dns"].append(server)

        tacacs = _metadata_value(
            md, ["tacacs_server_ip", "tacacs server ip", "tacacs_server"]
        )
        if tacacs:
            tacacs = str(ipaddress.ip_address(tacacs))
            if tacacs not in context["tacacs"]:
                context["tacacs"].append(tacacs)

        syslog = _metadata_value(
            md,
            ["syslog_server_ip", "rsyslog_server_ip", "syslog server ip", "syslog_server"],
        )
        if syslog:
            syslog = str(ipaddress.ip_address(syslog))
            if syslog not in context["syslog"]:
                context["syslog"].append(syslog)

    return context


def _get_urpf_command(md):
    """Optional LAN-side uRPF. Default OFF because DHCP/asymmetric designs need care."""
    value = _metadata_value(md, ["urpf_mode", "urpf mode", "uRPF mode"])
    if value is None:
        return None

    mode = value.strip().lower()
    if mode in {"", "off", "none", "disabled", "false", "0"}:
        return None
    if mode in {"strict", "rx"}:
        return "ip verify unicast source reachable-via rx"
    if mode in {"loose", "any"}:
        return "ip verify unicast source reachable-via any"
    raise ValueError("urpf_mode må være OFF, STRICT eller LOOSE")


def _parse_acl_ports(value, protocol):
    value = _normalise_cell_text(value).lower()
    if value in {"", "any", "*"}:
        return [None]

    if protocol not in {"tcp", "udp"}:
        raise ValueError(
            f"destination_port '{value}' kan bare brukes med TCP eller UDP"
        )

    allowed_names = {
        "ssh", "www", "http", "https", "domain", "dns", "tacacs", "syslog",
        "ntp", "bootps", "bootpc"
    }
    result = []
    tokens = value.replace(";", ",").replace(" ", ",").split(",")
    for token in (x.strip() for x in tokens):
        if not token:
            continue
        if "-" in token:
            start, end = token.split("-", 1)
            if not (start.isdigit() and end.isdigit()):
                raise ValueError(f"Ugyldig port-range '{token}'")
            start_i, end_i = int(start), int(end)
            if not (1 <= start_i <= end_i <= 65535):
                raise ValueError(f"Ugyldig port-range '{token}'")
            result.append(("range", str(start_i), str(end_i)))
        elif token.isdigit():
            port = int(token)
            if not 1 <= port <= 65535:
                raise ValueError(f"Ugyldig port '{token}'")
            result.append(("eq", str(port)))
        elif token in allowed_names:
            # IOS accepts common service names in extended ACLs.
            result.append(("eq", "domain" if token == "dns" else token))
        else:
            raise ValueError(f"Ukjent ACL-port/service '{token}'")

    return result or [None]


def _known_flow_vrfs(context):
    return set(context.get("vrf_networks", {})) | set(context.get("vrf_loopbacks", {}))


def _acl_destination_endpoints(
    destination,
    context,
    local_gateway=None,
    source_network=None,
):
    """Resolve one FLOW_POLICY destination.

    Normal service VRFs resolve to LAN/subnet destinations only. Router
    loopbacks are infrastructure addresses and are therefore excluded from
    normal VRF-to-VRF rules.

    MGMT is the exception: MGMT loopbacks are valid management-plane
    destinations (for example SSH to a router loopback), so destination=MGMT
    resolves to both routed MGMT LANs and MGMT loopbacks.

    If source and destination are the same VRF, the local source subnet is
    omitted because same-subnet host traffic never traverses this router ACL.
    """
    dest = destination.upper()

    if dest in {"ANY", "INTERNET"}:
        return ["any"]
    if dest == "DNS":
        values = context.get("dns", [])
        if not values:
            raise ValueError("FLOW_POLICY bruker DNS, men dns_servers er ikke satt i Excel")
        return [f"host {ip}" for ip in values]
    if dest == "TACACS":
        values = context.get("tacacs", [])
        if not values:
            raise ValueError("FLOW_POLICY bruker TACACS, men tacacs_server_ip mangler")
        return [f"host {ip}" for ip in values]
    if dest == "SYSLOG":
        values = context.get("syslog", [])
        if not values:
            raise ValueError("FLOW_POLICY bruker SYSLOG, men syslog_server_ip mangler")
        return [f"host {ip}" for ip in values]
    if dest == "SELF":
        if not local_gateway:
            raise ValueError("FLOW_POLICY destination=SELF krever lokal gateway")
        return [f"host {local_gateway}"]

    known = _known_flow_vrfs(context)
    if dest not in known:
        raise ValueError(
            f"Ukjent FLOW_POLICY-destination '{destination}'. "
            f"Kjente VRF-er: {', '.join(sorted(known))}; spesialverdier: "
            "INTERNET, ANY, DNS, DHCP, TACACS, SYSLOG, SELF"
        )

    source_network = str(source_network) if source_network is not None else None
    rendered = []
    for endpoint in context.get("vrf_networks", {}).get(dest, []):
        network_address = endpoint[1]
        wildcard = endpoint[2]
        if source_network is not None:
            candidate = ipaddress.ip_network(
                f"{network_address}/{wildcard}", strict=False
            )
            # wildcard is not a prefix mask, so construct from hostmask safely.
            mask_int = ((1 << 32) - 1) ^ int(ipaddress.IPv4Address(wildcard))
            candidate = ipaddress.ip_network(
                f"{network_address}/{ipaddress.IPv4Address(mask_int)}", strict=False
            )
            if str(candidate) == source_network:
                continue
        rendered.append(f"{network_address} {wildcard}")

    # MGMT loopbacks are intentional management-plane destinations. Keep them
    # addressable by whatever protocol/port the FLOW_POLICY row specifies.
    # Example: MGMT -> MGMT, tcp/22 permits SSH to remote router MGMT loopbacks.
    if dest == "MGMT":
        for endpoint in context.get("vrf_loopbacks", {}).get("MGMT", []):
            value = f"host {endpoint[1]}"
            if value not in rendered:
                rendered.append(value)

    return rendered


def _internet_internal_guards(context, source_vrf, source_network, local_gateway):
    """Return internal destinations that INTERNET must never mean.

    INTERNET is compiled as `any` only after these deny guards.  The guards are
    limited to the source VRF because those are the internal destinations that
    are normally routable from that VRF.  Explicit FLOW_POLICY exceptions are
    emitted before these guards, so an intentional internal permit still wins.
    """
    rendered = []

    if local_gateway:
        rendered.append(f"host {local_gateway}")

    source_network = str(source_network)
    for endpoint in context.get("vrf_networks", {}).get(source_vrf, []):
        network_address = endpoint[1]
        wildcard = endpoint[2]
        mask_int = ((1 << 32) - 1) ^ int(ipaddress.IPv4Address(wildcard))
        candidate = ipaddress.ip_network(
            f"{network_address}/{ipaddress.IPv4Address(mask_int)}", strict=False
        )
        if str(candidate) == source_network:
            continue
        value = f"{network_address} {wildcard}"
        if value not in rendered:
            rendered.append(value)

    for endpoint in context.get("vrf_loopbacks", {}).get(source_vrf, []):
        value = f"host {endpoint[1]}"
        if value not in rendered:
            rendered.append(value)

    return rendered


def _interface_keys_for_vrf(ip_data, intf_prefix, vrf_name):
    intf_nums = list(ip_data["interface"])
    keys = []
    rows = ip_data[ip_data["vrf"].astype(str).str.strip().str.upper() == vrf_name]
    for _, row in rows.iterrows():
        intf = row["interface"]
        vlan = row["vlan"]
        sub = intf_nums.count(intf) > 1
        if sub:
            keys.append(f"interface {intf_prefix}{intf}.{vlan}")
        else:
            keys.append(f"interface {intf_prefix}{intf}")
    return keys


def _insert_before_exit(lines, command):
    if command in lines:
        return
    try:
        idx = lines.index("exit")
    except ValueError:
        lines.append(command)
    else:
        lines.insert(idx, command)


def apply_flow_policy(config, network_info, ip_data, intf_prefix, policy, context):
    """Compile FLOW_POLICY to one ingress extended ACL per local source VRF.

    Specific internal/service rules are emitted before INTERNET rules. INTERNET
    then gets an automatic internal guard so it means "external destination",
    not simply "any routable address".
    """
    if policy is None or policy.empty:
        return

    valid_protocols = {"ip", "tcp", "udp", "icmp", "gre", "esp", "ah"}
    valid_actions = {"permit", "deny"}
    local_vrfs = {
        _normalise_cell_text(v).upper()
        for v in ip_data["vrf"].tolist()
        if _normalise_cell_text(v)
    }

    known_vrfs = _known_flow_vrfs(context)
    for source in policy["source"].tolist():
        source_u = _normalise_cell_text(source).upper()
        if source_u and source_u not in known_vrfs:
            raise ValueError(
                f"Ukjent FLOW_POLICY source '{source}'. Kjente VRF-er: "
                f"{', '.join(sorted(known_vrfs))}"
            )

    flow_info = {}
    for source in policy["source"].tolist():
        source_vrf = _normalise_cell_text(source).upper()
        if not source_vrf or source_vrf not in local_vrfs or source_vrf in flow_info:
            continue

        source_rows = policy[
            policy["source"].astype(str).str.strip().str.upper() == source_vrf
        ]
        ip_rows = ip_data[
            ip_data["vrf"].astype(str).str.strip().str.upper() == source_vrf
        ]
        if ip_rows.empty:
            continue

        source_row = ip_rows.iloc[0]
        source_net = ipaddress.ip_network(
            f"{source_row['nett id']}/{source_row['mask']}", strict=False
        )
        source_match = f"{source_net.network_address} {source_net.hostmask}"
        gateway = _normalise_cell_text(source_row.get("address min"))

        acl_name = f"FLOW-{source_vrf}-IN"
        acl_lines = []
        seen_lines = set()
        seen_rules = {}

        def append_line(line):
            if line not in seen_lines:
                acl_lines.append(line)
                seen_lines.add(line)

        # Validate rules and conflicts before reordering INTERNET behind specific
        # destinations. This keeps the Excel policy deterministic and catches
        # mistakes early.
        row_records = []
        for row_index, row in source_rows.iterrows():
            destination = _normalise_cell_text(row.get("destination")).upper()
            protocol = _normalise_cell_text(row.get("protocol")).lower()
            action = _normalise_cell_text(row.get("action")).lower()
            port_text = _normalise_cell_text(row.get("destination_port")).lower() or "any"

            if not destination:
                raise ValueError(f"FLOW_POLICY rad {row_index + 2}: destination mangler")
            if protocol not in valid_protocols:
                raise ValueError(
                    f"FLOW_POLICY rad {row_index + 2}: ugyldig protocol '{protocol}'"
                )
            if action not in valid_actions:
                raise ValueError(
                    f"FLOW_POLICY rad {row_index + 2}: action må være permit eller deny"
                )

            conflict_key = (source_vrf, destination, protocol, port_text)
            previous = seen_rules.get(conflict_key)
            if previous and previous != action:
                raise ValueError(
                    f"FLOW_POLICY har konflikt for {conflict_key}: både {previous} og {action}"
                )
            if previous == action:
                continue
            seen_rules[conflict_key] = action
            row_records.append((row_index, destination, protocol, action, port_text))

        specific_records = [r for r in row_records if r[1] != "INTERNET"]
        internet_records = [r for r in row_records if r[1] == "INTERNET"]

        def compile_record(record):
            row_index, destination, protocol, action, port_text = record

            if destination == "DHCP":
                if protocol != "udp":
                    raise ValueError("FLOW_POLICY destination=DHCP må bruke protocol=udp")
                if port_text not in {"67", "bootps", "any", "*", ""}:
                    raise ValueError("FLOW_POLICY destination=DHCP forventer destination_port=67")
                append_line(f"{action} udp any eq 68 any eq 67")
                return

            destinations = _acl_destination_endpoints(
                destination,
                context,
                local_gateway=gateway,
                source_network=source_net,
            )
            port_specs = _parse_acl_ports(port_text, protocol)

            # A same-VRF rule on a one-site workbook can legitimately resolve to
            # zero routed LAN destinations after the local subnet is removed.
            for dest_match in destinations:
                for spec in port_specs:
                    line = f"{action} {protocol} {source_match} {dest_match}"
                    if spec:
                        if spec[0] == "eq":
                            line += f" eq {spec[1]}"
                        else:
                            line += f" range {spec[1]} {spec[2]}"
                    append_line(line)

        # Explicit services and VRF-to-VRF exceptions come first.
        for record in specific_records:
            compile_record(record)

        # INTERNET is not synonymous with ANY. Before external permits, block
        # the router gateway, same-VRF remote LANs and same-VRF loopbacks. Any
        # explicit internal permit above remains effective because ACLs are
        # first-match.
        if any(record[3] == "permit" for record in internet_records):
            for internal_dest in _internet_internal_guards(
                context, source_vrf, source_net, gateway
            ):
                append_line(f"deny ip {source_match} {internal_dest}")

        for record in internet_records:
            compile_record(record)

        if not acl_lines:
            continue

        acl_lines.extend(["deny ip any any log", "exit"])
        config[f"ip access-list extended {acl_name}"] = acl_lines

        bound_interfaces = []
        for interface_key in _interface_keys_for_vrf(ip_data, intf_prefix, source_vrf):
            if interface_key not in config:
                raise ValueError(
                    f"Kunne ikke binde {acl_name}: finner ikke {interface_key} i generert config"
                )
            _insert_before_exit(config[interface_key], f"ip access-group {acl_name} in")
            bound_interfaces.append(interface_key.removeprefix("interface "))

        flow_info[source_vrf] = {
            "acl": acl_name,
            "interfaces": bound_interfaces,
            "rules": len(acl_lines) - 2,
            "default": "deny ip any any log",
        }

    if flow_info:
        network_info["flow_policy"] = flow_info

def create_vrf(vrf_data, sn):
    my_data = {}
    my_data["config"] = {}
    my_data["network_info"] = {}

    for index, row in vrf_data.iterrows():
        vrf_name = row["vrf"]
        vrf_loopback = row["loopback"]
        vrf_laddr = row["laddr"]

        my_data["network_info"][vrf_name] = {
            "loopback": vrf_loopback,
            "laddr": vrf_laddr
        }

        vrf_s = []
        vrf_s.append("exit")
        my_data["config"][f"ip vrf {vrf_name}"] = vrf_s

        intf_s = []
        intf_s.append(f"ip vrf forwarding {vrf_name}")
        intf_s.append(f"ip address {vrf_laddr} 255.255.255.255")
        intf_s.append("exit")
        my_data["config"][f"interface loopback{vrf_loopback}"] = intf_s

    return my_data


def create_interface(ip_data, intf_prefix, md=None):
    my_data = {}
    my_data["config"] = {}
    my_data["network_info"] = {}
    intf_nums = list(ip_data["interface"])
    
    tot_pri_num = ip_data["pri-1-10"].sum()
    max_prc = 75
    
    pol_maps= {}

    for index, row in ip_data.iterrows():
        vrf = row["vrf"]
        vlan = row["vlan"]
        pri_afxx = row["pri-afxx"]
        pri_num = row["pri-1-10"]
        top_num = str(pri_afxx)[0]

        if tot_pri_num == 0:
            pri_prc = 0
        else:
            pri_prc = int((pri_num / tot_pri_num)*max_prc)

        intf = row["interface"]
        sub = True if intf_nums.count(intf) > 1 else False

        ip_address = row["address min"]
        mask = row["mask"]

        if "interfaces" not in my_data["network_info"]:
            my_data["network_info"]["interfaces"] = {}

        my_data["network_info"]["interfaces"][
            f"{intf_prefix}{intf}.{vlan}" if sub else f"{intf_prefix}{intf}"
        ] = {
            "vrf": vrf,
            "vlan": vlan,
            "interface": intf,
            "sub": sub,
            "address": ip_address,
            "mask": mask,
            "pri_afxx": pri_afxx,
            "pri_num": pri_num,
        }

        intf_s = []

        if sub:
            intf_s.append(f"encapsulation dot1Q {vlan}")

            if f"interface {intf_prefix}{intf}.{999}" not in my_data["config"]:
                my_data["config"][
                    f"interface {intf_prefix}{intf}.{999}" if sub
                    else f"interface {intf_prefix}{intf}"
                ] = [
                    "description Sub-interface for ubrukt natiiv VLAN",
                    "encapsulation dot1Q 999 native",
                    "no ip address",
                    "no shutdown",
                    "exit"
                ]

        intf_s.append(f"ip vrf forwarding {vrf}")
        intf_s.append(f"ip address {ip_address} {mask}")
        intf_s.append("no ip redirects")
        intf_s.append("no ip proxy-arp")
        urpf_command = _get_urpf_command(md)
        if urpf_command:
            intf_s.append(urpf_command)
        intf_s.append("no shutdown")
        intf_s.append("exit")

        my_data["config"][
            f"interface {intf_prefix}{intf}.{vlan}" if sub
            else f"interface {intf_prefix}{intf}"
        ] = intf_s


        if sub:
            my_data["config"][f"class-map match-any QRS-MARK-{vrf}"] = [
                f"match vlan {vlan}",
                "exit"
            ]

            my_data["config"][f"class-map match-any QRS-{vrf}"] = [
                f"match dscp af{pri_afxx}",
                "exit"
            ]

            if f"policy-map QRS-SITE-MARK-POLICY" not in pol_maps:
                pol_maps[f"policy-map QRS-SITE-MARK-POLICY"] = []

            if f"policy-map QRS-SITE-POLICY" not in pol_maps:
                pol_maps[f"policy-map QRS-SITE-POLICY"] = []


            pol_maps[f"policy-map QRS-SITE-MARK-POLICY"].append(
                {
                    f"class QRS-MARK-{vrf}": [
                        f"set dscp af{pri_afxx}", 
                        "exit"
                    ]
                }   
            )

            pol_maps[f"policy-map QRS-SITE-POLICY"].append(
                {
                    f"class QRS-{vrf}": [
                        f"bandwidth percent {pri_prc}", 
                        "exit"
                    ]
                }   
            )
            

    intf_s = []
    intf_s.append("no shutdown")
    intf_s.append("exit")
    my_data["config"][f"interface {intf_prefix}{intf}"] = intf_s
    
    if f"policy-map QRS-SITE-MARK-POLICY" in pol_maps:
        my_data["config"].update(pol_maps)
        
        my_data["config"][f"\ninterface {intf_prefix}{intf}"] = [
            "service-policy input QRS-SITE-MARK-POLICY",
            "exit"
        ]

    if f"policy-map QRS-SITE-POLICY" in pol_maps:
        my_data["config"].update(pol_maps)
        
        my_data["config"][f"\ninterface {intf_prefix}1"] = [
            "service-policy output QRS-SITE-POLICY",
            "exit"
        ]

    return my_data


def create_ipsec_config(tunnel, source, sn, sites_data, network_id, vrf, is_hub,psk=DEFAULT_IPSEC_PSK):
    """
    Lager IPsec-konfigurasjon for en DMVPN-tunnel.
    """
    suffix = str(network_id).strip()

    proposal = f"DMVPN-IKEV2-PROP-{suffix}"
    policy = f"DMVPN-IKEV2-POL-{suffix}"
    keyring = f"DMVPN-IKEV2-KR-{suffix}"
    ikev2_profile = f"DMVPN-IKEV2-PROFILE-{suffix}"
    transform_set = f"DMVPN-TS-{suffix}"
    ipsec_profile = f"DMVPN-IPSEC-{suffix}"

    config = {}

    config[f"crypto ikev2 proposal {proposal}"] = [
        "encryption aes-cbc-256",
        "integrity sha256",
        "group 14",
        "exit",
    ]

    config[f"crypto ikev2 policy {policy}"] = [
        f"proposal {proposal}",
        "exit",
    ]

    remots_s = []
    my_stuff = source
    other_sites = sites_data.copy()
    del other_sites["hub"]

    if len(other_sites) > 0:
        if f"site {sn}" in other_sites:
            del other_sites[f"site {sn}"]

        for site, site_data in other_sites.items():
            tun = site_data["network_info"].get(tunnel, [])

            ip_address = tun.get("source", "")
            mask = "255.255.255.255"
            remots_s.append(f"address {ip_address} {mask}")

            # Oppdater allerede genererte sites når en ny DMVPN-peer blir kjent.
            # IKKE bruk set() her: Cisco CLI er rekkefølgeavhengig, og `exit` må stå sist.
            keyring_key = f"crypto ikev2 keyring {keyring}"
            profile_key = f"crypto ikev2 profile {ikev2_profile}"

            if keyring_key not in sites_data[site]["config"] or profile_key not in sites_data[site]["config"]:
                raise ValueError(
                    f"IPsec-oppsettet for {tunnel} er inkonsistent mellom site {sn} og {site}. "
                    "Samme DMVPN-cloud må bruke IPsec på alle deltakende sites."
                )

            site_pers = sites_data[site]["config"][keyring_key][0]["peer ANY"]
            site_pers = [f"address {my_stuff} {mask}"] + site_pers
            sites_data[site]["config"][keyring_key][0]["peer ANY"] = ensure_exit_last(site_pers)

            site_profile_pers = sites_data[site]["config"][profile_key]
            new_match = f"match identity remote address {my_stuff} {mask}"
            site_profile_pers = [new_match] + site_profile_pers
            sites_data[site]["config"][profile_key] = ensure_exit_last(site_profile_pers)


    peer = ensure_exit_last(
        remots_s
        + [
            f"pre-shared-key local {psk}",
            f"pre-shared-key remote {psk}",
        ]
    )

    config[f"crypto ikev2 keyring {keyring}"] = [
        {
            "peer ANY": peer
        },
        "exit",
    ]

    prof_peer = ensure_exit_last(
        [f"match identity remote {remote}" for remote in remots_s]
        + [
            "authentication remote pre-share",
            "authentication local pre-share",
            f"keyring local {keyring}",
        ]
    )

    config[f"crypto ikev2 profile {ikev2_profile}"] = prof_peer


    config[
        f"crypto ipsec transform-set {transform_set} esp-aes 256 esp-sha256-hmac"
    ] = [
        "mode transport",
        "exit",
    ]

    config[f"crypto ipsec profile {ipsec_profile}"] = [
        f"set transform-set {transform_set}",
        f"set ikev2-profile {ikev2_profile}",
        "exit",
    ]

    return config, ipsec_profile, sites_data


def create_tunnel_config(tunnel_data, sites_data: dict, is_hub: bool, sn):
    my_data = {}
    my_data["network_info"] = {}
    my_data["config"] = {}

    if is_hub:
        hub_data = {}
    else:
        hub = sites_data["hub"]
        hub_data = sites_data[hub]["network_info"]

    for idx, row in tunnel_data.iterrows():
        tunnel = row["tunnel id"]
        mode = row["gre mode"]
        ip_address = row["ip address"]
        mask = row["mask"]
        vrf = row["vrf"]
        source = row["source"]
        network_id = row["network-id"]

        ipsec_enabled = is_true(row.get("ipsec", False))

        # Valgfri egen PSK per tunnel.
        # om kollone tom brukes DEFAULT_IPSEC_PSK.
        psk = row.get("ipsec key", DEFAULT_IPSEC_PSK)
        if pd.isna(psk) or str(psk).strip() == "":
            psk = DEFAULT_IPSEC_PSK+f"-{network_id}"
        else:
            psk = str(psk).strip()

        if mode != "multipoint" and "destination" in row.index:
            destination = row["destination"]

        source = str(ipaddress.ip_address(source) + network_id+1)
        my_data["config"][f"interface loopback{network_id+1}"] = [
            f"description Loopback interface for tunnel/ipsec {tunnel}",
            f"ip address {source} 255.255.255.255",
            "ip ospf 1 area 0",
            "no shutdown",
            "exit",
        ]

        tunnel_info = {
            "is hub": is_hub,
            "vrf": vrf,
            "ip address": ip_address,
            "mask": mask,
            "source": source,
            "tunnel id": tunnel,
            "mode": mode,
            "ipsec": ipsec_enabled,
        }

        my_data["network_info"][tunnel] = tunnel_info

        tun_s = []
        tun_s.append(f"ip vrf forwarding {vrf}")
        tun_s.append(f"qos pre-classify")
        tun_s.append(f"ip address {ip_address} {mask}")
        tun_s.append(f"tunnel source loopback{network_id+1}")

        if mode == "multipoint":
            tun_s.append(f"tunnel mode gre {mode}")

            if not is_hub:
                hub_tun_ip = hub_data.get(tunnel, {}).get("ip address", "")
                hub_source = hub_data.get(tunnel, {}).get("source", "")

                tun_s.append(f"ip nhrp map {hub_tun_ip} {hub_source}")
                tun_s.append(f"ip nhrp map multicast {hub_source}")
                tun_s.append(f"ip nhrp nhs {hub_tun_ip}")
            else:
                tun_s.append("ip nhrp map multicast dynamic")

            tun_s.append(f"ip nhrp network-id {network_id}")
            if is_hub:
                tun_s.append("ip nhrp redirect")
            else:
                tun_s.append("ip nhrp shortcut")
            tun_s.append("no ip redirects")
            tun_s.append(f"tunnel key {network_id}")

        else:
            print(f"Mode må være multipoint for tunnel {tunnel}")

        # IPsec aktiveres bare når kolonnen 'ipsec' er TRUE/1/yes/ja/x.
        if ipsec_enabled:
            ipsec_config, ipsec_profile, sites_data = create_ipsec_config(tunnel,source, sn, sites_data, network_id, vrf, is_hub, psk)
            my_data["config"].update(ipsec_config)
            tun_s.append(f"tunnel protection ipsec profile {ipsec_profile}")

        tun_s.append("exit")

        my_data["config"][f"interface {tunnel}"] = tun_s
        my_data["network_info"][tunnel] = tunnel_info

    return my_data, sites_data


def create_tunnel_eigrp_config(vrf_data, tunnel_data, ip_data, is_hub):
    my_data = {}
    my_data["network_info"] = {}
    my_data["config"] = {}

    vrfs = {}
    networks = []

    for idx, row in tunnel_data.iterrows():
        network_id = row["network-id"]
        tun = row["tunnel id"]

        vrf = row["vrf"]
        vrfs[vrf] = []

        ip = row["ip address"]
        mask = row["mask"]

        netinfo = getNetId(ip, mask)
        networks.append(netinfo)

        vrfs[vrf].insert(0, network_id)
        vrfs[vrf].insert(1, tun)
        vrfs[vrf].append(netinfo)

    for idx, row in ip_data.iterrows():
        if row["vrf"] in vrfs:
            network = row["nett id"]
            mask = row["mask"]
            vrfs[row["vrf"]].append(getNetId(network, mask))

    tun_s = {}

    for vrf, nets in vrfs.items():
        tun_vrf_s = []

        for net in nets[2:]:
            if isinstance(net, tuple):
                network, mask, wild = net
                tun_vrf_s.append(f"network {network} {wild}")
           
        
        vrf_laddr = vrf_data[vrf_data["vrf"] == vrf].iloc[0].get("laddr", "")
        if vrf_laddr:
            tun_vrf_s.append(f"network {vrf_laddr} 0.0.0.0")

        tun_vrf_s.append(f"\n")
        tun_vrf_s.append(f"af-interface default")
        tun_vrf_s.append(f"passive-interface")
        tun_vrf_s.append(f"exit-af-interface\n")

        tun_vrf_s.append(f"af-interface {nets[1]}")
        if is_hub:
            tun_vrf_s.append("no split-horizon")
            tun_vrf_s.append("no next-hop-self")
        tun_vrf_s.append(f"no passive-interface")
        tun_vrf_s.append("exit-af-interface\n")

        if is_hub:
            if vrf == "INET":
                tun_vrf_s.append({"topology base": ["redistribute static metric 100000 10 255 1 1500", "exit-af-topology"]})

        tun_vrf_s.append("exit-address-family")
        tun_s[
            f"address-family ipv4 vrf {vrf} autonomous-system {nets[0]}"
        ] = tun_vrf_s

    tun_s["exit"] = []
    my_data["config"]["router eigrp DMVPN-EIGRP"] = tun_s

    return my_data


def _get_management_server_ip(md, names, label):
    """Read a management service IP directly from Excel top metadata."""
    if md is None or md.empty:
        raise ValueError(f"{label}-server mangler i Excel-metadata.")

    row = md.iloc[0]
    value = None
    for name in names:
        if name in row.index:
            candidate = row.get(name)
            if not pd.isna(candidate) and str(candidate).strip():
                value = candidate
                break

    if value is None:
        raise ValueError(f"{label}-server mangler i Excel. Forventet felt: {names[0]}.")

    try:
        return str(ipaddress.ip_address(str(value).strip()))
    except ValueError as exc:
        raise ValueError(f"Ugyldig {label}-server-IP i Excel: {value}") from exc


def create_tacacs_config(md, ip_data, sites_data, is_hub):
    my_data = {"config": {}, "network_info": {}}
    if md.empty:
        return my_data

    row = md.iloc[0]
    tacacs_server = _get_management_server_ip(
        md,
        ["tacacs_server_ip", "tacacs server ip", "tacacs_server"],
        "TACACS",
    )
    tacacs_key = row.get("tacacs_key", "")
    if pd.isna(tacacs_key) or not str(tacacs_key).strip():
        raise ValueError("TACACS-key mangler i Excel")

    if is_hub:
        print(f"TACACS server: {tacacs_server}")

    my_data["config"]["aaa new-model"] = []
    my_data["config"]["aaa group server tacacs+ TACACS-GROUP"] = [
        f"server-private {tacacs_server} key {str(tacacs_key).strip()}",
        "ip vrf forwarding MGMT",
        "ip tacacs source-interface loop10",
        "exit"
    ]
    my_data["config"]["aaa authentication login default group TACACS-GROUP local"] = []
    my_data["config"]["aaa authorization exec default group TACACS-GROUP local"] = []
    my_data["config"]["aaa accounting exec default start-stop group TACACS-GROUP"] = []
    my_data["config"]["aaa accounting commands 15 default start-stop group TACACS-GROUP"] = []
    my_data["config"]["line console 0"] = [
        "login authentication default",
        "exec-timeout 10 0",
        "logging synchronous",
        "exit",
    ]
    return my_data


def create_rsyslog_config(md, ip_data, sites_data, is_hub):
    my_data = {"config": {}, "network_info": {}}
    if md.empty:
        return my_data

    rsyslog_server = _get_management_server_ip(
        md,
        ["syslog_server_ip", "rsyslog_server_ip", "syslog server ip", "syslog_server"],
        "Syslog",
    )
    if is_hub:
        print(f"Rsyslog server: {rsyslog_server}")

    my_data["config"]["service timestamps log datetime msec show-timezone"] = []
    my_data["config"][f"logging host {rsyslog_server} vrf MGMT transport udp port 514"] = []
    my_data["config"]["logging trap informational"] = []
    my_data["config"]["logging buffered 16384 informational"] = []
    my_data["config"]["logging source-interface loop10 vrf MGMT"] = []
    return my_data


def enable_ssh(md, vrf_data, ip_data, sites_data, sn, domain=SSH_DOMAIN):
    """
    Genererer lokal SSH-konfig
    brukernavn + passord.
    """
    my_data = {}
    my_data["config"] = {}
    my_data["network_info"] = {}

    if md.empty:
        return my_data

    row = md.iloc[0]

    username = row.get("brukernavn", "")
    password = row.get("passord", "")
    vty_lines = row.get("vty_lines", "0-4")

    if pd.isna(username) or pd.isna(password):
        return my_data

    username = str(username).strip()
    password = str(password).strip()

    if not username or not password:
        return my_data

    my_data["config"][f"ip domain name {domain}"] = []
    my_data["config"][
        f"username {username} privilege 15 secret 9 {password}"
    ] = []
    my_data["config"]["crypto key generate rsa general-keys modulus 2048"] = []
    my_data["config"]["ip ssh version 2"] = []
    
    
    this_site = f"site {sn}"
    
    other_sites = sites_data.copy()

    if "hub" in other_sites:
        del other_sites["hub"]
        
    if this_site in other_sites:
        del other_sites[this_site]
    
    my_mgmt_ip_data = ip_data[ip_data["vrf"] == "MGMT"].iloc[0]
    my_network = my_mgmt_ip_data.get("nett id", "")
    my_mask = my_mgmt_ip_data.get("mask", "")
    my_network, my_mask, my_wild = getNetId(my_network, my_mask)
    
    my_vrf = vrf_data[vrf_data["vrf"] == "MGMT"].iloc[0]
    my_vrf_laddr = my_vrf.get("laddr", "")
    

    
    remots_s = []
    remots_s.append(f"permit {my_network} {my_wild}")
    remots_s.append(f"permit {my_vrf_laddr} 0.0.0.0")
    
    for site, site_data in other_sites.items():
        interfaces = site_data["network_info"].get("interfaces", [])
        for intf, intf_data in interfaces.items():
            vrf = intf_data.get("vrf", "")
            if vrf == "MGMT":
                ip_address = intf_data.get("address", "")
                network = str(ipaddress.ip_address(ip_address) - 1)
                mask = intf_data.get("mask", "")
                network, mask, wild = getNetId(network, mask)
                remots_s.append(f"permit {network} {wild}")

                if f"permit {my_network} {my_wild}" not in sites_data[site]["config"][f"ip access-list standard SSH-MGMT-ONLY"]:
                    sites_data[site]["config"][f"ip access-list standard SSH-MGMT-ONLY"].append(f"permit {my_network} {my_wild}")
                
    my_data["config"][f"ip access-list standard SSH-MGMT-ONLY"] = remots_s

    
    my_data["config"][f"line vty {' '.join(x.strip(' ') for x in vty_lines.split('-'))}"] = [
        "access-class SSH-MGMT-ONLY in vrf-also",
        "login authentication default",
        "exec-timeout 10 0",
        "transport input ssh",
        "exit",
    ]

    return my_data, sites_data


def fetch_site_data(config_file, site_number):
    try:
        data = {}
        hub = ""
        is_hub = False

        with open(config_file, "r") as f:
            try:
                data = json.load(f)
                hub = data.get("hub", "")
            except json.JSONDecodeError:
                data = {}
                data["hub"] = f"site {site_number}"
                hub = f"site {site_number}"

    except FileNotFoundError:
        data = {}
        data["hub"] = f"site {site_number}"
        hub = f"site {site_number}"

    return data, hub == f"site {site_number}"


def create_global_config(md, router_id, intf_prefix, sn, is_hub):
    my_data = {}
    my_data["config"] = {}
    my_data["network_info"] = {}

    secret = md.iloc[0].get("secret", "")

    my_data["config"][f"hostname RS{sn}"] = []
    my_data["config"][f"enable secret 9 {secret}"] = []
    my_data["config"]["service tcp-keepalives-in"] = []
    my_data["config"]["service tcp-keepalives-out"] = []
    my_data["config"]["no ip source-route"] = []
    my_data["config"]["banner motd ^CKun autorisert tilgang er tillatt. Aktivitet kan bli logget.^C"] = []

    my_data["config"]["interface loopback0"] = []
    my_data["config"]["interface loopback0"].append(
        f"ip address {router_id} 255.255.255.255"
    )
    my_data["config"]["interface loopback0"].append("ip ospf 1 area 0")
    my_data["config"]["interface loopback0"].append("exit")

    ospf_s = []
    ospf_s.append(f"router-id {router_id}")
    ospf_s.append("exit")
    my_data["config"]["router ospf 1"] = ospf_s


    dhcp = is_true(md.iloc[0].get("DHCP", False))
    dhcp_config = {}
    if is_hub and dhcp:
        dhcp_config = {
            "ip dhcp excluded-address 10.0.0.1 10.0.0.10": [],
            "ip dhcp pool CORE": [
                "network 10.0.0.0 255.255.255.0",
                "default-router 10.0.0.1",
                "exit"
            ]
        }
        
        intf_s = {
            f"interface {intf_prefix}1": [
                "ip address 10.0.0.1 255.255.255.0",
                "ip ospf 1 area 0",
                "no ip redirects",
                "no ip proxy-arp",
                "no shutdown",
                "exit"
            ]
        }
        
        for key, value in dhcp_config.items():
            my_data["config"][key] = value
        for key, value in intf_s.items():
            my_data["config"][key] = value
    else:
        intf_s = []
        intf_s.append("ip address dhcp")
        intf_s.append("ip ospf 1 area 0")
        intf_s.append("no ip redirects")
        intf_s.append("no ip proxy-arp")
        intf_s.append("no shutdown")
        intf_s.append("exit")
        my_data["config"][f"interface {intf_prefix}1"] = intf_s
    


    my_data["network_info"]["loopback0"] = {
        "address": router_id,
        "mask": "255.255.255.255"
    }

    return my_data


def get_dns_servers(md):
    """Return validated DNS server IPs from the optional Excel dns_servers field."""
    if md.empty or "dns_servers" not in md.columns:
        return []

    value = md.iloc[0].get("dns_servers", "")
    if pd.isna(value) or not str(value).strip():
        return []

    # Accept spaces, commas, or semicolons in the Excel cell.
    raw = str(value).replace(",", " ").replace(";", " ")
    servers = [item.strip() for item in raw.split() if item.strip()]
    for server in servers:
        try:
            ipaddress.ip_address(server)
        except ValueError as exc:
            raise ValueError(f"Ugyldig DNS-server i Excel: {server}") from exc
    return servers


def set_up_DHCP_for_vrf_lans(ip_data, md):
    my_config = {}
    my_config["config"] = {}
    my_config["network_info"] = {}

    dns_servers = get_dns_servers(md)

    for idx, row in ip_data.iterrows():
        vrf = row["vrf"]
        ip_gw = row["address min"]
        network = row["nett id"]
        mask = row["mask"]
        num_res = int(float(row["antall-res"]))

        ip_res_to = str(ipaddress.ip_address(ip_gw) + num_res)

        pool = [
            f"vrf {vrf}",
            f"network {network} {mask}",
            f"default-router {ip_gw}",
        ]
        if vrf == "INET" and dns_servers:
            pool.append(f"dns-server {' '.join(dns_servers)}")
        pool.append("exit")

        my_config["config"][f"ip dhcp pool DHCP-{vrf}"] = pool

        my_config["config"][f"ip dhcp excluded-address vrf {vrf} {ip_gw} {ip_res_to}"] = []

        my_config["network_info"][f"DHCP-{vrf}"] = {
            "network": network,
            "mask": mask,
            "default-router": ip_gw,
            "ip_res_to": ip_res_to
        }
        
    return my_config


def config_nat(md, is_hub):
    my_config = {}
    my_config["config"] = {}
    my_config["network_info"] = {}

    if is_hub:
        intf_prefix = md.iloc[0]["intf_prefix"]

        my_config["config"]["!\ninterface tunnel20"] = [
            "ip nat inside",
            "exit"
        ]

        my_config["config"][f"interface {intf_prefix}0.200"] = [
            "encapsulation dot1Q 200",
            "ip address dhcp",
            "ip nat outside",
            "no shutdown",
            "exit"
        ]
        my_config["config"]["ip access-list standard NAT-INET"] = [
            "permit any",
            "exit"
        ]
        my_config["config"][f"ip nat inside source list NAT-INET interface {intf_prefix}0.200 vrf INET overload"] = []
        my_config["config"][f"!\ninterface {intf_prefix}0.20"] = [
            "ip nat inside",
            "exit"
        ]

        my_config["config"][f"ip route vrf INET 0.0.0.0 0.0.0.0 {intf_prefix}0.200 dhcp"] = []

    return my_config


def config_ntp(sites_data, is_hub):
    my_config = {}
    my_config["config"] = {}
    my_config["network_info"] = {}


    if is_hub:
        my_config["config"]["ntp master 8"] = []
    else:
        hub_site = sites_data["hub"]
        hub_info = sites_data[hub_site]
        server_ip = hub_info["network_info"]["loopback0"]["address"]

        my_config["config"][f"ntp server {server_ip}"] = []

    return my_config
    
    
def configure_site(sheet_file, config_file, sheet, flow_policy=None, flow_context=None):
    sheet_data = read_sheet(sheet_file, sheet)

    my_data = {}
    my_data["config"] = {}
    my_data["network_info"] = {}

    #HENT NØDVENDIG DATA 
    ip_data = sheet_data["ip_data"]
    vrf_data = sheet_data["vrf_data"]
    tunnel_data = sheet_data["tunnel_data"]
    md = sheet_data["md"]

    router_id = md.iloc[0]["router-id"]
    sn = md.iloc[0]["site"]
    intf_prefix = md.iloc[0]["intf_prefix"]

    data, is_hub = fetch_site_data(config_file, sn)
    
    
    #OPPRETT GLOBAL KONFIGURASJON
    my_data = create_global_config(md,router_id, intf_prefix, sn, is_hub)

    #NTP
    d_ntp = config_ntp(data, is_hub)
    my_data["config"].update(d_ntp["config"])
    
    #VRF
    d_vrf = create_vrf(vrf_data, sn)
    my_data["config"].update(d_vrf["config"])
    my_data["network_info"].update(d_vrf["network_info"])

    #TACACS
    d_tacacs = create_tacacs_config(md, ip_data, data, is_hub)
    my_data["config"].update(d_tacacs["config"])

    #RSYSLOG
    d_rsyslog = create_rsyslog_config(md, ip_data, data, is_hub)
    my_data["config"].update(d_rsyslog["config"])



    #INTERFACE
    d_ip = create_interface(ip_data, intf_prefix, md)
    my_data["config"].update(d_ip["config"])
    my_data["network_info"].update(d_ip["network_info"])

    # FLOW POLICY / DATA-PLANE ACL
    if flow_policy is not None and not flow_policy.empty:
        apply_flow_policy(
            my_data["config"],
            my_data["network_info"],
            ip_data,
            intf_prefix,
            flow_policy,
            flow_context or {},
        )
    
    #DHCP
    d_dhcp = set_up_DHCP_for_vrf_lans(ip_data, md)
    my_data["config"].update(d_dhcp["config"])
    my_data["network_info"].update(d_dhcp["network_info"])

    #TUNNEL
    d_tunnel, data = create_tunnel_config(tunnel_data, data, is_hub, sn)
    my_data["config"].update(d_tunnel["config"])
    my_data["network_info"].update(d_tunnel["network_info"])

    #EIGRP
    d_eigrp = create_tunnel_eigrp_config(vrf_data, tunnel_data, ip_data, is_hub)
    my_data["config"].update(d_eigrp["config"])
    my_data["network_info"].update(d_eigrp["network_info"])

    #NAT
    d_nat = config_nat(md, is_hub)
    my_data["config"].update(d_nat["config"])

    #SSH
    d_ssh, data = enable_ssh(md, vrf_data, ip_data, data, sn)
    my_data["config"].update(d_ssh["config"])   
     
    #INTerFACE PREFIX
    my_data["intf_prefix"] = intf_prefix

    #LAGRE
    data[f"site {sn}"] = my_data

    with open(config_file, "w") as f:
        json.dump(data, f, indent=4)

    return data


def config_to_text(data, indent=0):
    lines = []
    prefix = "    " * indent

    if isinstance(data, dict):
        for key, value in data.items():
            lines.append("!")
            lines.append(prefix + key)
            lines.extend(config_to_text(value, indent + 1))

    elif isinstance(data, list):
        for value in data:
            if isinstance(value, str):
                lines.append(prefix + value)
            else:
                lines.extend(config_to_text(value, indent))

    elif isinstance(data, str):
        lines.append(prefix + data)
        
    lines.append("!")
  
    return lines


def create_or_update_config_files(data):
    del data["hub"]

    if not os.path.exists("siteEdgeRouterTextConfigs"):
        os.makedirs("siteEdgeRouterTextConfigs")

    for site, site_data in data.items():
        config = site_data["config"]
        text = config_to_text(config)

        with open(
            f"siteEdgeRouterTextConfigs/EDGE_ROUTER_{site.replace(' ', '_').upper()}.txt",
            "w",
            encoding="utf-8"
        ) as f:
            f.write("\n".join(text))
    
    print()
    print(f"Text versjon av config for edge router i site {site} er fullført og lagret i siteEdgeRouterTextConfigs/EDGE_ROUTER_{site.replace(' ', '_').upper()}.txt")
    print()
    

def create_edge_router_configs_main(file, config_file="EDGE_ROUTER_configs.json"):   

    sites_sheets = _site_sheet_names(file)
    flow_policy = read_flow_policy(file)
    flow_context = (
        build_flow_policy_context(file, sites_sheets)
        if flow_policy is not None and not flow_policy.empty
        else {}
    )
    
    for sheet in sites_sheets:
        data = configure_site(
            file, config_file, sheet, flow_policy=flow_policy, flow_context=flow_context
        )
    
    create_or_update_config_files(data)


def main():
    if not os.path.exists("siteEdgeRouterTextConfigs"):
        os.makedirs("siteEdgeRouterTextConfigs")

    sheet_file = sys.argv[1]
    config_file = (
        sys.argv[2]
        if len(sys.argv) > 2
        else "EDGE_ROUTER_configs.json"
    )

    create_edge_router_configs_main(sheet_file, config_file)


if __name__ == "__main__":
    main()