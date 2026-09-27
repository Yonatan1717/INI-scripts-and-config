import json
import ipaddress
import os
import re
import sys
import shutil
from pathlib import Path, PurePosixPath

BASE_DIR = Path(__file__).resolve().parent
NETWORK_DEV_SCRIPTS = BASE_DIR / "networkDevScripts"
if str(NETWORK_DEV_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(NETWORK_DEV_SCRIPTS))

from site_EDGE_ROUTER_script import create_edge_router_configs_main
from site_SWITCH_script import create_sw_configs_main


# Ansible inventory-innstillinger.
# Passord legges med vilje ikke i inventory-filen. Bruk f.eks. Ansible Vault.
ANSIBLE_USER = "admin"
ANSIBLE_PASSWORD = "bani"
LEGACY_SSH = True  # Sett False dersom enhetene støtter moderne SSH-algoritmer.
LEGACY_SSH_CONFIG_NAME = "legacy_ssh.cfg"
LEGACY_SSH_CIPHERS = "+aes128-cbc"
LEGACY_SSH_MACS = "+hmac-sha1"


# Midlertidig staging-/INIT-switch for førstegangsoppsett.
# Generatoren lager flere interface-varianter slik at man kan velge den som
# passer den fysiske svitsjen som brukes som konfigurasjonssvitsj.
INIT_SWITCH_MGMT_VLAN = 10
INIT_SWITCH_NATIVE_VLAN = 999
INIT_SWITCH_CONTROL_NODE_PORT = 1

# Vanlige Cisco interface-navn. Alle variantene genereres automatisk.
INIT_SWITCH_INTERFACE_VARIANTS = {
    "GI0": "GigabitEthernet0/",
    "GI1_0": "GigabitEthernet1/0/",
    "FA0": "FastEthernet0/",
}

# Control-node-adressene under er midlertidige adresser på samme fysiske NIC.
# Dagens MGMT-design bruker /24 per site. Generatoren velger høyeste ledige
# hostadresse blant 254..240 som ikke kolliderer med genererte Cisco-enheter.
INIT_CONTROL_NODE_PREFIXLEN = 24
INIT_CONTROL_NODE_FIRST_CANDIDATE = 254
INIT_CONTROL_NODE_LAST_CANDIDATE = 240


def _load_json(path):
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _hostname_from_config(config, fallback=None):
    """Hent hostname fra en generert Cisco-config-dict."""
    for command in config:
        command_clean = str(command).strip()
        if command_clean.lower().startswith("hostname "):
            return command_clean.split(None, 1)[1].strip()
    return fallback


def _static_ip_from_interface_commands(commands):
    """Returner statisk IP fra en liste med interface-kommandoer."""
    if not isinstance(commands, list):
        return None

    for command in commands:
        if not isinstance(command, str):
            continue

        parts = command.strip().split()
        if len(parts) >= 3 and parts[0:2] == ["ip", "address"]:
            if parts[2].lower() != "dhcp":
                return parts[2]

    return None


def _router_mgmt_ip(site_data):
    """Hent MGMT-adressen til en site-router fra network_info."""
    network_info = site_data.get("network_info", {})
    interfaces = network_info.get("interfaces", {})

    for interface_data in interfaces.values():
        if str(interface_data.get("vrf", "")).upper() == "MGMT":
            address = interface_data.get("address")
            if address:
                return str(address)

    # Fallback dersom JSON-strukturen endres senere.
    config = site_data.get("config", {})
    for commands in config.values():
        if not isinstance(commands, list):
            continue

        has_mgmt_vrf = any(
            isinstance(command, str)
            and command.strip().lower() == "ip vrf forwarding mgmt"
            for command in commands
        )
        if has_mgmt_vrf:
            address = _static_ip_from_interface_commands(commands)
            if address:
                return address

    return None


def _switch_mgmt_ip(config):
    """Hent switchens management-SVI. VLAN 10 prioriteres, ellers første SVI med statisk IP."""
    candidates = []

    for interface_name, commands in config.items():
        name = str(interface_name).strip().lower()
        if not name.startswith("interface vlan"):
            continue

        address = _static_ip_from_interface_commands(commands)
        if not address:
            continue

        vlan_part = name.removeprefix("interface vlan").strip()
        try:
            vlan_id = int(vlan_part)
        except ValueError:
            vlan_id = 99999

        candidates.append((vlan_id, address))

    if not candidates:
        return None

    # MGMT er VLAN 10 i dagens design. Hvis den finnes velges den alltid.
    for vlan_id, address in candidates:
        if vlan_id == 10:
            return address

    return sorted(candidates, key=lambda item: item[0])[0][1]


def _collect_router_inventory(router_data):
    routers = []

    for site_name, site_data in router_data.items():
        if site_name == "hub" or not isinstance(site_data, dict):
            continue

        config = site_data.get("config", {})
        hostname = _hostname_from_config(config)
        mgmt_ip = _router_mgmt_ip(site_data)

        if not hostname:
            raise ValueError(f"Fant ikke hostname for router i {site_name}")
        if not mgmt_ip:
            raise ValueError(f"Fant ikke MGMT-IP for router {hostname} i {site_name}")

        routers.append((hostname, mgmt_ip, site_name))

    return routers


def _collect_switch_inventory(switch_data):
    switches = []

    for site_name, site_data in switch_data.items():
        if str(site_name).startswith("_") or not isinstance(site_data, dict):
            continue

        devices = site_data.get("config", {})
        for device_name, config in devices.items():
            if not isinstance(config, dict):
                continue

            hostname = _hostname_from_config(config, fallback=device_name)
            mgmt_ip = _switch_mgmt_ip(config)

            if not mgmt_ip:
                raise ValueError(f"Fant ikke MGMT-IP for switch {hostname} i {site_name}")

            switches.append((hostname, mgmt_ip, site_name))

    return switches


def _generate_legacy_ssh_config(routers, switches, output_dir):
    """Lag en OpenSSH-config som kun aktiverer legacy cipher/MAC for genererte Cisco-enheter."""
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    config_path = output_dir / LEGACY_SSH_CONFIG_NAME

    # Behold samme rekkefølge som inventory, men fjern eventuelle duplikate IP-er.
    device_ips = []
    seen = set()
    for _hostname, ip, _site in routers + switches:
        ip = str(ip).strip()
        if ip and ip not in seen:
            seen.add(ip)
            device_ips.append(ip)

    if not device_ips:
        raise ValueError("Kan ikke generere legacy SSH-config uten Cisco management-IP-er.")

    ssh_lines = [
        "# Generated automatically by ultimate_config_script_site_router_and_switch.py",
        "# Legacy SSH algorithms are enabled ONLY for the Cisco management IPs below.",
        "",
        f"Host {' '.join(device_ips)}",
        f"    Ciphers {LEGACY_SSH_CIPHERS}",
        f"    MACs {LEGACY_SSH_MACS}",
        "",
    ]

    config_path.write_text("\n".join(ssh_lines), encoding="utf-8")
    print(f"Legacy SSH-config generert: {config_path}")
    return config_path




def _site_slug(site_name):
    return str(site_name).strip().lower().replace(" ", "_")


def _copy_ansible_config_trees(output_dir, bundle_dir):
    """Flytt genererte Ansible-configtrær til prosjektets ansible_folder.

    Router-/switchgeneratorene skriver først de Ansible-klare filene under
    networkConfigs mens de kjører. Hovedscriptet flytter dem deretter inn i
    ansible_folder, slik at all Ansible-data ender samlet på ett sted.
    """
    bundle_dir.mkdir(parents=True, exist_ok=True)
    for dirname in ("ansibleConfigs", "ansibleBootstrapConfigs"):
        source = output_dir / dirname
        destination = bundle_dir / dirname
        if not source.exists():
            raise FileNotFoundError(f"Fant ikke generert Ansible-mappe: {source}")
        if destination.exists():
            shutil.rmtree(destination)
        shutil.move(str(source), str(destination))


def _inventory_device_line(hostname, ip, site_name, bundle_dir):
    """Build one host line using paths relative to inventory.ini.

    This deliberately avoids absolute Windows/Linux paths.  The whole Ansible
    bundle can therefore be copied between Windows, WSL and Linux unchanged.
    """
    site_slug = _site_slug(site_name)
    full_rel = PurePosixPath("ansibleConfigs") / site_slug / f"{hostname}.cfg"
    bootstrap_rel = (
        PurePosixPath("ansibleBootstrapConfigs")
        / site_slug
        / f"{hostname}_SSH_SCP.cfg"
    )

    full_path = bundle_dir.joinpath(*full_rel.parts)
    bootstrap_path = bundle_dir.joinpath(*bootstrap_rel.parts)
    if not full_path.exists():
        raise FileNotFoundError(f"Fant ikke Ansible-ready config for {hostname}: {full_path}")
    if not bootstrap_path.exists():
        raise FileNotFoundError(f"Fant ikke SSH/SCP-bootstrap for {hostname}: {bootstrap_path}")

    return (
        f'{hostname} ansible_host={ip} '
        f'config_file="{full_rel.as_posix()}" '
        f'bootstrap_config_file="{bootstrap_rel.as_posix()}"'
    )


def _suggest_control_node_addresses(routers, switches):
    """Lag én midlertidig control-node-IP per MGMT-/24.

    Alle staging-portene ligger i VLAN 10, men hvert site bruker sitt eget
    IP-subnett. Linux-control-noden må derfor ha en adresse i hvert subnett
    under staging uten å være avhengig av ferdig ruting.
    """
    used_by_network = {}

    for _hostname, ip, _site in routers + switches:
        addr = ipaddress.ip_address(str(ip))
        network = ipaddress.ip_network(
            f"{addr}/{INIT_CONTROL_NODE_PREFIXLEN}", strict=False
        )
        used_by_network.setdefault(network, set()).add(addr)

    suggestions = []
    for network in sorted(used_by_network, key=lambda n: int(n.network_address)):
        used = used_by_network[network]
        chosen = None

        for host_octet in range(
            INIT_CONTROL_NODE_FIRST_CANDIDATE,
            INIT_CONTROL_NODE_LAST_CANDIDATE - 1,
            -1,
        ):
            candidate = network.network_address + host_octet
            if (
                candidate in network
                and candidate not in used
                and candidate != network.broadcast_address
            ):
                chosen = candidate
                break

        if chosen is None:
            raise ValueError(
                f"Fant ingen ledig foreslått control-node-adresse i {network} "
                f"mellom .{INIT_CONTROL_NODE_FIRST_CANDIDATE} og "
                f".{INIT_CONTROL_NODE_LAST_CANDIDATE}."
            )

        suggestions.append((network, chosen))

    return suggestions


def _init_switch_config(routers, switches, interface_prefix):
    """Bygg en midlertidig L2 staging-switch-konfigurasjon.

    Port 1 er control node. Routere får trunkporter fordi router-bootstrapen
    bruker dot1Q-subinterface for MGMT VLAN 10. Switcher får accessporter i
    MGMT VLAN 10.
    """
    lines = [
        "!",
        "! ================================================================",
        "! TEMPORARY INIT / CONFIG SWITCH",
        "! Brukes kun til førstegangsoppsett før full Ansible-deploy.",
        "! Port 1 = Ansible control node",
        "! Router-porter = trunk med tagged MGMT VLAN 10",
        "! Switch-porter = access i MGMT VLAN 10",
        "! ================================================================",
        "!",
        "hostname INIT-CONFIG-SW",
        "vtp mode transparent",
        "!",
        f"vlan {INIT_SWITCH_MGMT_VLAN}",
        " name MGMT-STAGING",
        "!",
        f"vlan {INIT_SWITCH_NATIVE_VLAN}",
        " name NATIVE-BLACKHOLE",
        "!",
        f"interface {interface_prefix}{INIT_SWITCH_CONTROL_NODE_PORT}",
        " description ANSIBLE-CONTROL-NODE",
        " switchport mode access",
        f" switchport access vlan {INIT_SWITCH_MGMT_VLAN}",
        " spanning-tree portfast",
        " no shutdown",
        "!",
    ]

    port = INIT_SWITCH_CONTROL_NODE_PORT + 1
    port_map = [
        (
            INIT_SWITCH_CONTROL_NODE_PORT,
            "CONTROL-NODE",
            "ACCESS",
            f"VLAN {INIT_SWITCH_MGMT_VLAN}",
        )
    ]

    # Routere må få tagged VLAN 10 fordi MGMT ligger på dot1Q-subinterface.
    for hostname, ip, site_name in routers:
        lines.extend(
            [
                f"interface {interface_prefix}{port}",
                f" description STAGING-{hostname}-{ip}",
                f" switchport trunk native vlan {INIT_SWITCH_NATIVE_VLAN}",
                " switchport mode trunk",
                " switchport nonegotiate",
                f" switchport trunk allowed vlan {INIT_SWITCH_MGMT_VLAN},{INIT_SWITCH_NATIVE_VLAN}",
                " spanning-tree portfast trunk",
                " no shutdown",
                "!",
            ]
        )
        port_map.append(
            (
                port,
                hostname,
                "TRUNK",
                f"tagged VLAN {INIT_SWITCH_MGMT_VLAN} ({ip}, {site_name})",
            )
        )
        port += 1

    # Switch-management-SVI mottar vanlig untagged access VLAN 10.
    for hostname, ip, site_name in switches:
        lines.extend(
            [
                f"interface {interface_prefix}{port}",
                f" description STAGING-{hostname}-{ip}",
                " switchport mode access",
                f" switchport access vlan {INIT_SWITCH_MGMT_VLAN}",
                " spanning-tree portfast",
                " no shutdown",
                "!",
            ]
        )
        port_map.append(
            (
                port,
                hostname,
                "ACCESS",
                f"VLAN {INIT_SWITCH_MGMT_VLAN} ({ip}, {site_name})",
            )
        )
        port += 1

    lines.extend(
        [
            "end",
            "",
            "! MERK:",
            "! Denne switchen er kun et midlertidig staging-L2-nett.",
            "! Ikke koble staging-VLANet til produksjonsnettet samtidig.",
            "",
        ]
    )
    return "\n".join(lines), port_map


def _generate_init_switch_bundle(routers, switches, bundle_dir):
    """Generer staging-switch, portkart og control-node nettverksmal."""
    init_dir = Path(bundle_dir) / "init_config_switch"
    init_dir.mkdir(parents=True, exist_ok=True)

    port_map = None
    for variant, interface_prefix in INIT_SWITCH_INTERFACE_VARIANTS.items():
        config_text, this_port_map = _init_switch_config(
            routers, switches, interface_prefix
        )
        cfg_path = init_dir / f"INIT_CONFIG_SWITCH_{variant}.cfg"
        cfg_path.write_text(config_text, encoding="utf-8")
        if port_map is None:
            port_map = this_port_map

    map_lines = [
        "INIT / CONFIG SWITCH - PORT MAP",
        "================================",
        "",
        f"MGMT VLAN: {INIT_SWITCH_MGMT_VLAN}",
        f"Native/blackhole VLAN: {INIT_SWITCH_NATIVE_VLAN}",
        "",
        "Port | Rolle/enhet                | Type   | Formål",
        "-----+----------------------------+--------+----------------------------------------------",
    ]
    for port, device, mode, purpose in port_map or []:
        map_lines.append(
            f"{port:>4} | {device:<26} | {mode:<6} | {purpose}"
        )
    map_lines.extend(
        [
            "",
            "Velg config som matcher interface-navnet på staging-switchen:",
            "  GI0   = GigabitEthernet0/x",
            "  GI1_0 = GigabitEthernet1/0/x",
            "  FA0   = FastEthernet0/x",
            "",
            "Routerportene er trunk fordi router-bootstrapen bruker:",
            "  encapsulation dot1Q 10",
            "",
            "Switchporter og control-node-port er access i VLAN 10.",
        ]
    )
    (init_dir / "PORT_MAP.txt").write_text(
        "\n".join(map_lines) + "\n", encoding="utf-8"
    )

    suggestions = _suggest_control_node_addresses(routers, switches)

    shell_lines = [
        "#!/usr/bin/env bash",
        "# Midlertidige adresser for Ansible-control-node under staging.",
        "# BYTT <ANSIBLE_NIC> med riktig interface, f.eks. ens160.",
        "# Kontroller at foreslåtte adresser er ledige før de tas i bruk.",
        "",
        'NIC="${1-NONE}"',
        'if [ "$NIC" = "NONE" ]; then',
        '    echo "Usage: $0 <ANSIBLE_NIC>"',
        '    exit 1',
        'fi',
        "",
    ]
    for network, address in suggestions:
        shell_lines.append(
            f'sudo ip addr add {address}/{network.prefixlen} dev "$NIC"'
        )

    shell_lines.extend(
        [
            "",
            "# Verifisering:",
            'ip -br addr show "$NIC"',
            "",
            "# Fjern adressene etter staging dersom de ikke skal beholdes:",
        ]
    )
    for network, address in suggestions:
        shell_lines.append(
            f'# sudo ip addr del {address}/{network.prefixlen} dev "$NIC"'
        )

    (init_dir / "CONTROL_NODE_TEMP_IPS.sh").write_text(
        "\n".join(shell_lines) + "\n", encoding="utf-8"
    )

    readme_lines = [
        "INIT / CONFIG SWITCH",
        "====================",
        "",
        "Formål:",
        "  Lage ett midlertidig L2-stagingnett der Ansible-control-noden kan nå",
        "  alle Cisco-enhetene etter at deres *_SSH_SCP.cfg er lagt inn via console.",
        "",
        "Fysisk oppkobling:",
        "  Port 1         -> Ansible control node (ACCESS VLAN 10)",
        "  Neste porter   -> Site-routere (TRUNK, tagged VLAN 10)",
        "  Resterende     -> Site-switcher (ACCESS VLAN 10)",
        "",
        "Hvorfor router-portene er trunk:",
        "  Router-bootstrapen legger MGMT-adressen på et dot1Q-subinterface.",
        "  VLAN 10 må derfor komme tagged inn til routeren.",
        "",
        "VIKTIG om control node:",
        "  Site 1, Site 2, Site 3 osv. bruker ulike MGMT-IP-subnett selv om alle",
        "  bruker VLAN 10. INIT-switchen ruter ikke mellom dem. Control-noden må",
        "  derfor ha én midlertidig IP-adresse i hvert MGMT-subnett på samme NIC.",
        "  CONTROL_NODE_TEMP_IPS.sh inneholder forslag til slike adresser.",
        "",
        "Arbeidsflyt:",
        "  1. Velg riktig INIT_CONFIG_SWITCH_*.cfg og konfigurer staging-switchen.",
        "  2. Koble Ansible-serveren til port 1.",
        "  3. Legg midlertidige site-MGMT-adresser på Ansible-NIC-en.",
        "  4. Koble routere til trunkportene og switcher til accessportene.",
        "  5. Lim inn riktig *_SSH_SCP.cfg via console på hver target-enhet.",
        "  6. Verifiser SSH fra Ansible-serveren.",
        "  7. Kjør deploy_generated.yml.",
        "  8. Verifiser og kjør write memory.",
        "  9. Flytt ferdig konfigurerte enheter til endelig topologi.",
        "",
        "Dette er et midlertidig stagingnett og skal ikke stå parallelt koblet mot",
        "produksjonsnettet under førstegangsoppsettet.",
    ]
    (init_dir / "README_INIT_SWITCH.txt").write_text(
        "\n".join(readme_lines) + "\n", encoding="utf-8"
    )

    print(f"INIT/config-switch bundle generert: {init_dir}")
    return init_dir


def generate_ansible_inventory(output_dir, store_ini_in):
    """Create a self-contained, portable Ansible bundle.

    All per-device config paths in inventory.ini are relative to inventory_dir,
    so the bundle may be moved/copied without rewriting Windows/Linux paths.
    """
    output_dir = Path(output_dir).resolve()
    store_ini_in = Path(store_ini_in).resolve()
    store_ini_in.mkdir(parents=True, exist_ok=True)
    _copy_ansible_config_trees(output_dir, store_ini_in)

    router_json = output_dir / "EDGE_ROUTER_configs.json"
    switch_json = output_dir / "site_switch_config.json"

    if not router_json.exists():
        raise FileNotFoundError(f"Fant ikke router-JSON: {router_json}")
    if not switch_json.exists():
        raise FileNotFoundError(f"Fant ikke switch-JSON: {switch_json}")

    router_data = _load_json(router_json)
    switch_data = _load_json(switch_json)

    routers = _collect_router_inventory(router_data)
    switches = _collect_switch_inventory(switch_data)

    _generate_init_switch_bundle(routers, switches, store_ini_in)

    # Oppdag duplikate hostnames før vi skriver en ugyldig inventory.
    all_hosts = [hostname for hostname, _ip, _site in routers + switches]
    duplicate_hosts = sorted({h for h in all_hosts if all_hosts.count(h) > 1})
    if duplicate_hosts:
        raise ValueError(
            "Duplikate hostnames i Ansible inventory: " + ", ".join(duplicate_hosts)
        )

    lines = ["[cisco_routers]"]
    lines.extend(
        _inventory_device_line(hostname, ip, site_name, store_ini_in)
        for hostname, ip, site_name in routers
    )

    lines.extend(["", "[cisco_switches]"])
    lines.extend(
        _inventory_device_line(hostname, ip, site_name, store_ini_in)
        for hostname, ip, site_name in switches
    )

    lines.extend(
        [
            "",
            "[cisco:children]",
            "cisco_routers",
            "cisco_switches",
            "",
            "[cisco:vars]",
            "ansible_connection=ansible.netcommon.network_cli",
            "ansible_network_os=cisco.ios.ios",
            f"ansible_user={ANSIBLE_USER}",
            f"ansible_password={ANSIBLE_PASSWORD}",
            "ansible_network_cli_ssh_type=libssh",
            "ansible_host_key_checking=False",
        ]
    )

    legacy_ssh_path = store_ini_in / LEGACY_SSH_CONFIG_NAME

    if LEGACY_SSH:
        # OpenSSH-configen brukes til cipher/MAC som de gamle Cisco-enhetene krever.
        # KEX/hostkey beholdes som libssh-variabler siden dette allerede fungerer i laben.
        legacy_ssh_path = _generate_legacy_ssh_config(routers, switches, store_ini_in)
        lines.extend(
            [
                "ansible_libssh_config_file=\"{{ inventory_dir }}/legacy_ssh.cfg\"",
                "ansible_libssh_key_exchange_algorithms=+diffie-hellman-group14-sha1",
                "ansible_libssh_hostkeys=ssh-rsa",
            ]
        )
    elif legacy_ssh_path.exists():
        # Unngå at en gammel legacy-config blir liggende igjen når funksjonen skrus av.
        legacy_ssh_path.unlink()

    inventory_path = store_ini_in / "inventory.ini"
    inventory_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    deploy_path = store_ini_in / "deploy_generated.yml"
    deploy_path.write_text(
        "---\n"
        "- name: Deploy generated Cisco configuration\n"
        "  hosts: cisco\n"
        "  gather_facts: false\n"
        "  connection: ansible.netcommon.network_cli\n"
        "  serial: 1\n\n"
        "  tasks:\n"
        "    - name: Apply generated configuration and save if changed\n"
        "      cisco.ios.ios_config:\n"
        "        src: \"{{ inventory_dir }}/{{ config_file }}\"\n"
        "        backup: true\n"
        "        save_when: modified\n",
        encoding="utf-8",
    )

    readme_path = store_ini_in / "README_GENERATED_CONFIGS.txt"
    readme_path.write_text(
        "Generated Ansible files\n"
        "=======================\n\n"
        "This directory is a self-contained portable Ansible bundle.\n"
        "The paths in inventory.ini are RELATIVE paths, not Windows/Linux absolute paths.\n"
        "You can therefore copy the whole directory to the same or another host.\n\n"
        "inventory.ini contains two per-device variables:\n"
        "  config_file           = relative full Ansible-ready configuration\n"
        "  bootstrap_config_file = relative console-paste MGMT + SSH/SCP bootstrap\n\n"
        "init_config_switch/ contains the temporary staging-switch configs, port map,\n"
        "and control-node temporary MGMT IP instructions used before Ansible deploy.\n\n"
        "Bootstrap is intended to make the device reachable by Ansible first.\n"
        "It uses local VTY authentication and does not depend on TACACS/RADIUS.\n"
        "After SSH works, a playbook can apply the full file with:\n\n"
        "  cisco.ios.ios_config:\n"
        "    src: '{{ inventory_dir }}/{{ config_file }}'\n"
        "    backup: true\n"
        "    save_when: modified\n\n"
        "Important: ios_config src performs a merge. A command that disappears\n"
        "from the generated file is not automatically negated on the device.\n",
        encoding="utf-8",
    )

    print(f"Ansible inventory generert: {inventory_path}")
    print(f"Deploy-playbook generert: {deploy_path}")
    print(f"Ansible config-veiledning generert: {readme_path}")
    return inventory_path


def main():
    if len(sys.argv) < 2:
        raise SystemExit(
            "Bruk: python ultimate_config_script_site_router_and_switch.py <excel-fil>"
        )

    excel_file = Path(sys.argv[1]).resolve()
    if not excel_file.exists():
        raise FileNotFoundError(f"Fant ikke Excel-filen: {excel_file}")

    # Fast prosjektstruktur:
    # NO_MPLS_SAD/
    #   ansible_folder/
    #   networkConfigs/
    #   networkDevScripts/
    #   ultimate_config_script_site_router_and_switch.py
    #
    # All Ansible-relatert output samles i den eksisterende ansible_folder.
    ansible_bundle_dir = BASE_DIR / "ansible_folder"
    ansible_bundle_dir.mkdir(parents=True, exist_ok=True)

    print(f"Ansible-data lagres i prosjektmappen: {ansible_bundle_dir}")

    output_dir = BASE_DIR / "networkConfigs"
    output_dir.mkdir(exist_ok=True)

    org_cwd = Path.cwd()
    os.chdir(output_dir)
    try:
        create_edge_router_configs_main(excel_file)
        create_sw_configs_main(excel_file)
        generate_ansible_inventory(output_dir, ansible_bundle_dir)
    finally:
        os.chdir(org_cwd)


if __name__ == "__main__":
    main()
