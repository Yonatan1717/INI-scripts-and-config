#!/usr/bin/env python3
"""
summarize_site_config.py

Leser allerede genererte JSON-filer fra nettverksgeneratoren og skriver en
brukervennlig oppsummering av en eller alle sites. Støtter både MPLS og
NO-MPLS, inkludert SPAN/RSPAN/ERSPAN, MONITORING-SVI, ERSPAN source-ranges,
IP Source Guard og eventuell dedikert Suricata/IDS-sensorport.

Krever:
  - EDGE_ROUTER_configs.json
  - site_switch_config.json

Hvis JSON-filene ikke finnes eller er ugyldige, avsluttes programmet med feil
fordi det da ikke finnes en ferdig konfigurasjon å oppsummere.

Eksempler:
    python summarize_site_config.py
    python summarize_site_config.py 1
    python summarize_site_config.py --site 2
    python summarize_site_config.py --save summary.txt
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

ROUTER_JSON_NAME = "EDGE_ROUTER_configs.json"
SWITCH_JSON_NAME = "site_switch_config.json"


def fail(message: str, code: int = 1) -> None:
    print(f"FEIL: {message}", file=sys.stderr)
    raise SystemExit(code)


def natural_key(value: str) -> list[Any]:
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r"(\d+)", str(value))]


def find_default_json(filename: str) -> Path | None:
    script_dir = Path(__file__).resolve().parent
    cwd = Path.cwd()
    candidates = [
        cwd / filename,
        cwd / "networkConfigs" / filename,
        script_dir / filename,
        script_dir / "networkConfigs" / filename,
    ]

    seen = set()
    for candidate in candidates:
        candidate = candidate.resolve()
        if candidate in seen:
            continue
        seen.add(candidate)
        if candidate.is_file():
            return candidate
    return None


def resolve_json_path(explicit: str | None, filename: str) -> Path:
    if explicit:
        path = Path(explicit).expanduser().resolve()
        if not path.is_file():
            fail(f"Fant ikke {filename}: {path}")
        return path

    found = find_default_json(filename)
    if found is None:
        fail(
            f"Fant ikke {filename}. Kjør config-generatoren først. "
            f"Scriptet leter i current directory og ./networkConfigs/."
        )
    return found


def load_json(path: Path) -> dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as exc:
        fail(f"{path.name} finnes, men inneholder ugyldig JSON: {exc}")
    except OSError as exc:
        fail(f"Kunne ikke lese {path}: {exc}")

    if not isinstance(data, dict):
        fail(f"{path.name} har uventet format; forventet et JSON-objekt.")
    return data


def site_number(site_key: str) -> str:
    match = re.search(r"(\d+)", str(site_key))
    return match.group(1) if match else str(site_key).strip()


def normalize_site_arg(value: str) -> str:
    value = str(value).strip()
    if value.lower().startswith("site"):
        return f"site {site_number(value)}"
    if value.isdigit():
        return f"site {int(value)}"
    return value


def site_keys(data: dict[str, Any]) -> list[str]:
    return sorted(
        [k for k in data if str(k).lower().startswith("site ")],
        key=natural_key,
    )


def config_has_key(config: dict[str, Any], prefix: str) -> bool:
    prefix_l = prefix.lower()
    return any(str(k).lower().startswith(prefix_l) for k in config)


def first_config_key(config: dict[str, Any], prefix: str) -> str | None:
    prefix_l = prefix.lower()
    for key in config:
        if str(key).lower().startswith(prefix_l):
            return str(key)
    return None


def get_hostname(config: dict[str, Any], fallback: str) -> str:
    key = first_config_key(config, "hostname ")
    return key.split(maxsplit=1)[1] if key else fallback


def get_line_value(lines: Any, prefix: str) -> str | None:
    if not isinstance(lines, list):
        return None
    prefix_l = prefix.lower()
    for line in lines:
        if isinstance(line, str) and line.strip().lower().startswith(prefix_l):
            return line.strip()[len(prefix):].strip()
    return None


def get_line_values(lines: Any, prefix: str) -> list[str]:
    """Return all matching config lines, without the prefix."""
    if not isinstance(lines, list):
        return []
    prefix_l = prefix.lower()
    values = []
    for line in lines:
        if isinstance(line, str) and line.strip().lower().startswith(prefix_l):
            values.append(line.strip()[len(prefix):].strip())
    return values


def detect_router_technologies(config: dict[str, Any], network_info: dict[str, Any]) -> list[str]:
    tech = []

    if config_has_key(config, "router ospf "):
        tech.append("OSPF")

    if config_has_key(config, "mpls ldp router-id") or any(
        isinstance(v, list) and any(isinstance(x, str) and x.strip() == "mpls ip" for x in v)
        for v in config.values()
    ):
        tech.append("MPLS/LDP")

    if config_has_key(config, "router bgp "):
        tech.append("MP-BGP")

    if any(str(k).lower().startswith("tunnel") for k in network_info) or config_has_key(
        config, "router eigrp dmvpn-eigrp"
    ):
        tech.append("DMVPN/EIGRP")

    if config_has_key(config, "crypto ikev2 profile "):
        tech.append("IKEv2/IPsec")

    if config_has_key(config, "ip nat inside source"):
        tech.append("NAT")

    return tech


def get_router_services(config: dict[str, Any]) -> list[str]:
    services = []
    if "aaa new-model" in config:
        services.append("TACACS+/AAA")
    if config_has_key(config, "logging host "):
        services.append("Syslog")
    if "ip ssh version 2" in config:
        services.append("SSH")
    if config_has_key(config, "ntp server ") or config_has_key(config, "ntp master"):
        services.append("NTP")
    if any(str(k).startswith("ip dhcp pool ") for k in config):
        services.append("DHCP")
    return services


def extract_dns_from_pool(config: dict[str, Any], pool_name: str) -> list[str]:
    lines = config.get(f"ip dhcp pool {pool_name}", [])
    if not isinstance(lines, list):
        return []
    for line in lines:
        if isinstance(line, str) and line.strip().startswith("dns-server "):
            return line.strip().split()[1:]
    return []


def summarize_router(site_key: str, site_data: dict[str, Any], router_root: dict[str, Any]) -> list[str]:
    config = site_data.get("config", {})
    network_info = site_data.get("network_info", {})
    hostname = get_hostname(config, f"Router {site_number(site_key)}")
    role = "HUB" if router_root.get("hub") == site_key else "SPOKE"

    lines = [f"Router: {hostname} ({role})"]

    loop0 = network_info.get("loopback0", {})
    if isinstance(loop0, dict) and loop0.get("address"):
        lines.append(f"  Router-ID / Loopback0: {loop0['address']}")

    technologies = detect_router_technologies(config, network_info)
    if technologies:
        lines.append(f"  Teknologi: {', '.join(technologies)}")

    services = get_router_services(config)
    if services:
        lines.append(f"  Tjenester: {', '.join(services)}")

    vrfs = []
    for name, info in network_info.items():
        if isinstance(info, dict) and "loopback" in info and "laddr" in info:
            vrfs.append((str(name), info))

    if vrfs:
        lines.append("  VRF-er:")
        for vrf, info in sorted(vrfs, key=lambda x: natural_key(x[0])):
            rt = f", RT {info['rt']}" if info.get("rt") else ""
            lines.append(
                f"    - {vrf}: Loopback{info.get('loopback')} = {info.get('laddr')}{rt}"
            )

    interfaces = network_info.get("interfaces", {})
    if isinstance(interfaces, dict) and interfaces:
        lines.append("  LAN/VLAN:")
        entries = []
        for intf, info in interfaces.items():
            if not isinstance(info, dict):
                continue
            vlan = info.get("vlan")
            try:
                vlan_sort = int(vlan)
            except (TypeError, ValueError):
                vlan_sort = 999999
            entries.append((vlan_sort, str(intf), info))

        for _, intf, info in sorted(entries):
            qos = ""
            if info.get("pri_afxx") not in (None, ""):
                qos = f", QoS AF{info.get('pri_afxx')}"
            lines.append(
                f"    - VLAN {info.get('vlan')} / {info.get('vrf')}: "
                f"{intf} = {info.get('address')} {info.get('mask')}{qos}"
            )

    dhcp_entries = []
    for name, info in network_info.items():
        if str(name).startswith("DHCP-") and isinstance(info, dict):
            dhcp_entries.append((str(name), info))

    if dhcp_entries:
        lines.append("  DHCP:")
        for pool, info in sorted(dhcp_entries, key=lambda x: natural_key(x[0])):
            vrf = pool.removeprefix("DHCP-")
            dns = extract_dns_from_pool(config, pool)
            dns_text = f", DNS {' '.join(dns)}" if dns else ""
            lines.append(
                f"    - {vrf}: {info.get('network')} {info.get('mask')}, "
                f"GW {info.get('default-router')}, reservert til {info.get('ip_res_to')}{dns_text}"
            )

    flow_policy = network_info.get("flow_policy", {})
    if isinstance(flow_policy, dict) and flow_policy:
        lines.append("  Trafikkpolicy (extended ingress ACL):")
        for source, info in sorted(flow_policy.items(), key=lambda x: natural_key(x[0])):
            if not isinstance(info, dict):
                continue
            interfaces = ", ".join(info.get("interfaces", [])) or "ukjent interface"
            lines.append(
                f"    - {source}: {info.get('acl')} på {interfaces} "
                f"({info.get('rules', '?')} regler + default deny)"
            )

    tunnels = []
    for name, info in network_info.items():
        if str(name).lower().startswith("tunnel") and isinstance(info, dict):
            tunnels.append((str(name), info))

    if tunnels:
        lines.append("  Tunneler:")
        for name, info in sorted(tunnels, key=lambda x: natural_key(x[0])):
            ipsec = "IPsec" if info.get("ipsec") else "uten IPsec"
            tunnel_cfg = config.get(f"interface {name}", [])
            phase3 = ""
            if isinstance(tunnel_cfg, list):
                if "ip nhrp redirect" in tunnel_cfg:
                    phase3 = ", Phase 3 HUB/redirect"
                elif "ip nhrp shortcut" in tunnel_cfg:
                    phase3 = ", Phase 3 SPOKE/shortcut"
            lines.append(
                f"    - {name}: {info.get('vrf')} {info.get('ip address')} "
                f"({info.get('mode')}), source {info.get('source')}, {ipsec}{phase3}"
            )

    hardening = []
    if "no ip source-route" in config:
        hardening.append("IP source-route av")
    if config_has_key(config, "aaa accounting commands 15 "):
        hardening.append("AAA accounting")
    if config_has_key(config, "logging buffered "):
        hardening.append("buffered logging")
    if any(
        isinstance(v, list) and "no ip redirects" in v
        for v in config.values()
    ):
        hardening.append("ICMP redirects av")
    if any(
        isinstance(v, list) and "no ip proxy-arp" in v
        for v in config.values()
    ):
        hardening.append("Proxy ARP av")
    if any(
        isinstance(v, list)
        and any(isinstance(x, str) and x.startswith("ip verify unicast source reachable-via") for x in v)
        for v in config.values()
    ):
        hardening.append("uRPF")
    if hardening:
        lines.append(f"  Hardening: {', '.join(hardening)}")

    return lines


def switch_vlans(config: dict[str, Any]) -> list[int]:
    vlans = []
    for key in config:
        match = re.fullmatch(r"vlan\s+(\d+)", str(key).strip(), re.I)
        if match:
            vlans.append(int(match.group(1)))
    return sorted(set(vlans))


def summarize_switch(sw_name: str, config: dict[str, Any]) -> list[str]:
    lines = [f"  {sw_name}:"]

    svi_info: dict[str, str] = {}
    for key, value in config.items():
        match = re.fullmatch(r"interface\s+vlan\s+(\d+)", str(key).strip(), re.I)
        if not match:
            continue
        ip = get_line_value(value, "ip address ")
        if ip:
            svi_info[match.group(1)] = ip

    access_entries = []
    mgmt_ports = []
    mgmt_port_vlan = None
    span_port = None
    erspan_sensor_port = None
    uplinks = []
    downlinks = []
    port_channels = []

    for key, cfg in config.items():
        if not str(key).lower().startswith("interface") or not isinstance(cfg, list):
            continue

        description = get_line_value(cfg, "description ") or ""
        desc_l = description.lower()
        access_vlan = get_line_value(cfg, "switchport access vlan ")
        display_if = re.sub(r"^interface\s+", "", str(key), flags=re.I)

        if (
            "dedicated management" in desc_l
            or "dedicated tacacs" in desc_l
            or "dedicated syslog" in desc_l
        ):
            mgmt_ports.append((display_if, description, access_vlan))
            if mgmt_port_vlan is None and access_vlan:
                mgmt_port_vlan = access_vlan

        if "span destination" in desc_l or "rspan destination" in desc_l:
            span_port = display_if

        if "security onion" in desc_l and "mirror destination" in desc_l:
            erspan_sensor_port = display_if

        if (
            access_vlan
            and "access port for vlan" in desc_l
            and "dedicated management" not in desc_l
        ):
            access_entries.append((display_if, access_vlan))

        if "uplink" in desc_l and "port-channel" not in str(key).lower():
            uplinks.append(
                (display_if, description, get_line_value(cfg, "switchport trunk allowed vlan "))
            )
        elif "downlink" in desc_l and "port-channel" not in str(key).lower():
            downlinks.append(
                (display_if, description, get_line_value(cfg, "switchport trunk allowed vlan "))
            )

        if str(key).lower().startswith("interface port-channel"):
            port_channels.append(
                (display_if, get_line_value(cfg, "switchport trunk allowed vlan "))
            )

    mgmt_vlan = mgmt_port_vlan
    if not mgmt_vlan:
        for key, value in config.items():
            match = re.fullmatch(r"vlan\s+(\d+)", str(key).strip(), re.I)
            if not match or not isinstance(value, list):
                continue
            if any(isinstance(line, str) and "MGMT_VLAN_" in line.upper() for line in value):
                mgmt_vlan = match.group(1)
                break

    if mgmt_vlan and mgmt_vlan in svi_info:
        lines.append(f"    MGMT SVI: VLAN {mgmt_vlan} = {svi_info[mgmt_vlan]}")
    elif svi_info:
        vlan, ip = sorted(svi_info.items(), key=lambda x: int(x[0]))[0]
        lines.append(f"    MGMT SVI: VLAN {vlan} = {ip}")

    erspan_source_key = next(
        (str(k) for k in config if str(k).lower().startswith("monitor session ") and "type erspan-source" in str(k).lower()),
        None,
    )
    erspan_destination_key = next(
        (str(k) for k in config if str(k).lower().startswith("monitor session ") and "type erspan-destination" in str(k).lower()),
        None,
    )
    erspan_present = bool(erspan_source_key or erspan_destination_key)

    monitoring_vlan = None
    if erspan_present:
        for vlan in sorted(svi_info, key=int):
            if vlan != mgmt_vlan:
                monitoring_vlan = vlan
                break
    if monitoring_vlan and monitoring_vlan in svi_info:
        lines.append(f"    MONITORING SVI: VLAN {monitoring_vlan} = {svi_info[monitoring_vlan]}")

    gateway_key = first_config_key(config, "ip default-gateway ")
    if gateway_key:
        lines.append(f"    Default gateway: {gateway_key.split()[-1]}")
    else:
        default_route = first_config_key(config, "ip route 0.0.0.0 0.0.0.0 ")
        if default_route:
            lines.append(f"    Default route: {' '.join(default_route.split()[5:])}")

    vlans = switch_vlans(config)
    if vlans:
        lines.append(f"    VLANs: {', '.join(map(str, vlans))}")

    for port, description, vlan in mgmt_ports:
        suffix = f" | VLAN {vlan}" if vlan else ""
        lines.append(f"    {description}: {port}{suffix}")

    if erspan_sensor_port:
        lines.append(f"    Security Onion sniff/mirror-port: {erspan_sensor_port}")

    rspan_source = next(
        (str(k) for k in config if str(k).lower().startswith("monitor session ") and "source remote vlan" in str(k).lower()),
        None,
    )
    rspan_destination = next(
        (str(k) for k in config if str(k).lower().startswith("monitor session ") and "destination remote vlan" in str(k).lower()),
        None,
    )

    if erspan_destination_key:
        cfg = config.get(erspan_destination_key, [])
        dst_ip = get_line_value(cfg, "ip address ") or "ukjent"
        erspan_id = get_line_value(cfg, "erspan-id ") or "ukjent"
        destination_interface = get_line_value(cfg, "destination interface ") or erspan_sensor_port or "ukjent"
        lines.append(
            f"    Overvåking: ERSPAN destination | IP {dst_ip} | ERSPAN-ID {erspan_id} | "
            f"ut {destination_interface} til Security Onion"
        )
    elif erspan_source_key:
        cfg = config.get(erspan_source_key, [])
        dst = get_line_value(cfg, "ip address ") or "ukjent"
        erspan_id = get_line_value(cfg, "erspan-id ") or "ukjent"
        origin = get_line_value(cfg, "origin ip-address ") or "ukjent"
        source_interfaces = get_line_values(cfg, "source interface ")

        lines.append(
            f"    Overvåking: ERSPAN source | global routing | origin {origin} | "
            f"HUB destination {dst} | ERSPAN-ID {erspan_id}"
        )
        if source_interfaces:
            lines.append("    ERSPAN-kilder:")
            for source in source_interfaces:
                lines.append(f"      - {source}")

        route_prefix = f"ip route {dst} 255.255.255.255 "
        erspan_route = first_config_key(config, route_prefix)
        if erspan_route:
            next_hop = erspan_route[len(route_prefix):].strip()
            lines.append(f"    ERSPAN-rute: {dst}/32 via MONITORING-gateway {next_hop}")
    elif rspan_source:
        vlan = rspan_source.lower().split("source remote vlan", 1)[1].strip()
        port_text = f" | IDS-port {span_port}" if span_port else ""
        lines.append(f"    Overvåking: RSPAN destination | VLAN {vlan}{port_text}")
    elif rspan_destination:
        vlan = rspan_destination.lower().split("destination remote vlan", 1)[1].strip()
        lines.append(f"    Overvåking: RSPAN source | VLAN {vlan}")
    elif span_port:
        source = first_config_key(config, "monitor session 1 source ")
        source_text = source.removeprefix("monitor session 1 source ") if source else "ukjent"
        lines.append(f"    Overvåking: SPAN | IDS-port {span_port} | kilde {source_text}")

    if access_entries:
        lines.append("    Accessporter:")
        for port, vlan in access_entries:
            lines.append(f"      - {port}: VLAN {vlan}")

    if uplinks:
        lines.append("    Uplink:")
        for port, description, allowed in uplinks:
            suffix = f" | VLAN {allowed}" if allowed else ""
            lines.append(f"      - {port}: {description}{suffix}")

    if downlinks:
        lines.append("    Downlinks:")
        for port, description, allowed in downlinks:
            suffix = f" | VLAN {allowed}" if allowed else ""
            lines.append(f"      - {port}: {description}{suffix}")

    if port_channels:
        lines.append("    Port-channels:")
        for po, allowed in port_channels:
            suffix = f" | VLAN {allowed}" if allowed else ""
            lines.append(f"      - {po}{suffix}")

    security = []
    if "ip dhcp snooping" in config:
        security.append("DHCP snooping")
    if any(str(k).lower().startswith("ip arp inspection vlan ") for k in config):
        security.append("DAI")
    if any(
        isinstance(v, list)
        and any(isinstance(line, str) and line.strip() == "ip verify source" for line in v)
        for v in config.values()
    ):
        security.append("IP Source Guard")
    if any(isinstance(v, list) and "switchport port-security" in v for v in config.values()):
        security.append("Port-security")
    if any(isinstance(v, list) and "spanning-tree bpduguard enable" in v for v in config.values()):
        security.append("BPDU Guard")
    if any(isinstance(v, list) and "spanning-tree guard root" in v for v in config.values()):
        security.append("Root Guard")
    if any(isinstance(v, list) and "spanning-tree guard loop" in v for v in config.values()):
        security.append("Loop Guard")
    if any(isinstance(v, list) and "switchport nonegotiate" in v for v in config.values()):
        security.append("DTP av")
    if "vtp mode transparent" in config:
        security.append("VTP transparent")
    if "aaa new-model" in config:
        security.append("TACACS+/AAA")
    if config_has_key(config, "aaa accounting commands 15 "):
        security.append("AAA accounting")
    if "ip ssh version 2" in config:
        security.append("SSH")
    if config_has_key(config, "logging host "):
        security.append("Syslog")

    if security:
        lines.append(f"    Sikkerhet/drift: {', '.join(security)}")

    return lines


def summarize_site(
    site_key: str,
    router_data: dict[str, Any],
    switch_data: dict[str, Any],
) -> str:
    if site_key not in router_data:
        fail(f"{site_key} finnes ikke i {ROUTER_JSON_NAME}.")
    if site_key not in switch_data:
        fail(f"{site_key} finnes ikke i {SWITCH_JSON_NAME}.")

    lines = [
        "=" * 72,
        f"SITE {site_number(site_key)}",
        "=" * 72,
    ]
    lines.extend(summarize_router(site_key, router_data[site_key], router_data))

    switches = switch_data[site_key].get("config", {})
    lines.append("")
    lines.append(f"Switcher: {len(switches)}")
    for sw_name in sorted(switches, key=natural_key):
        lines.extend(summarize_switch(sw_name, switches[sw_name]))

    return "\n".join(lines)


def build_summary(
    router_data: dict[str, Any],
    switch_data: dict[str, Any],
    requested_site: str | None,
) -> str:
    router_sites = set(site_keys(router_data))
    switch_sites = set(site_keys(switch_data))
    common_sites = sorted(router_sites & switch_sites, key=natural_key)

    if not common_sites:
        fail("Fant ingen sites som finnes i både router- og switch-JSON.")

    if requested_site:
        site_key = normalize_site_arg(requested_site)
        if site_key not in common_sites:
            available = ", ".join(site_number(x) for x in common_sites)
            fail(f"Fant ikke {site_key} i begge JSON-filene. Tilgjengelige sites: {available}")
        sites = [site_key]
    else:
        sites = common_sites

    header = [
        "NETTVERKSKONFIGURASJON - OPPSUMMERING",
        f"Sites: {', '.join(site_number(x) for x in sites)}",
        "Merk: passord, enable secrets og TACACS-nøkler vises ikke.",
        "",
    ]

    return "\n\n".join(
        ["\n".join(header)]
        + [summarize_site(site, router_data, switch_data) for site in sites]
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Oppsummer allerede genererte router- og switch-konfigurasjoner."
    )
    parser.add_argument(
        "site_positional",
        nargs="?",
        help="Valgfritt site-nummer, f.eks. 1. Uten verdi oppsummeres alle sites.",
    )
    parser.add_argument("--site", dest="site_option", help="Valgfritt site-nummer.")
    parser.add_argument("--router-json", help=f"Sti til {ROUTER_JSON_NAME}.")
    parser.add_argument("--switch-json", help=f"Sti til {SWITCH_JSON_NAME}.")
    parser.add_argument(
        "--save",
        metavar="FIL",
        help="Lagre samme oppsummering til tekstfil i tillegg til å skrive den ut.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.site_positional and args.site_option:
        fail("Bruk enten posisjonelt site-nummer eller --site, ikke begge.")

    requested_site = args.site_option or args.site_positional

    router_path = resolve_json_path(args.router_json, ROUTER_JSON_NAME)
    switch_path = resolve_json_path(args.switch_json, SWITCH_JSON_NAME)

    router_data = load_json(router_path)
    switch_data = load_json(switch_path)

    summary = build_summary(router_data, switch_data, requested_site)
    print(summary)

    if args.save:
        output = Path(args.save).expanduser()
        try:
            output.write_text(summary + "\n", encoding="utf-8")
        except OSError as exc:
            fail(f"Kunne ikke lagre oppsummeringen til {output}: {exc}")
        print(f"\nOppsummering lagret til: {output}")


if __name__ == "__main__":
    main()