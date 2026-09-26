import json
import os
import sys
from pathlib import Path

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

        routers.append((hostname, mgmt_ip))

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

            switches.append((hostname, mgmt_ip))

    return switches


def _generate_legacy_ssh_config(routers, switches, output_dir):
    """Lag en OpenSSH-config som kun aktiverer legacy cipher/MAC for genererte Cisco-enheter."""
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    config_path = output_dir / LEGACY_SSH_CONFIG_NAME

    # Behold samme rekkefølge som inventory, men fjern eventuelle duplikate IP-er.
    device_ips = []
    seen = set()
    for _hostname, ip in routers + switches:
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


def generate_ansible_inventory(output_dir, store_ini_in):
    """Lag inventory.ini og eventuell legacy_ssh.cfg fra genererte router/switch-JSON-filer."""
    output_dir = Path(output_dir).resolve()
    store_ini_in = Path(store_ini_in).resolve()
    store_ini_in.mkdir(parents=True, exist_ok=True)

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

    # Oppdag duplikate hostnames før vi skriver en ugyldig inventory.
    all_hosts = [hostname for hostname, _ in routers + switches]
    duplicate_hosts = sorted({h for h in all_hosts if all_hosts.count(h) > 1})
    if duplicate_hosts:
        raise ValueError(
            "Duplikate hostnames i Ansible inventory: " + ", ".join(duplicate_hosts)
        )

    lines = ["[cisco_routers]"]
    lines.extend(f"{hostname} ansible_host={ip}" for hostname, ip in routers)

    lines.extend(["", "[cisco_switches]"])
    lines.extend(f"{hostname} ansible_host={ip}" for hostname, ip in switches)

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
                f"ansible_libssh_config_file={legacy_ssh_path}",
                "ansible_libssh_key_exchange_algorithms=+diffie-hellman-group14-sha1",
                "ansible_libssh_hostkeys=ssh-rsa",
            ]
        )
    elif legacy_ssh_path.exists():
        # Unngå at en gammel legacy-config blir liggende igjen når funksjonen skrus av.
        legacy_ssh_path.unlink()

    inventory_path = store_ini_in / "inventory.ini"
    inventory_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    print(f"Ansible inventory generert: {inventory_path}")
    return inventory_path


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Bruk: python ultimate_config_script_site_router_and_switch.py <excel-fil>")

    excel_file = Path(sys.argv[1]).resolve()
    if not excel_file.exists():
        raise FileNotFoundError(f"Fant ikke Excel-filen: {excel_file}")

    output_dir = BASE_DIR / "networkConfigs"
    output_dir.mkdir(exist_ok=True)

    org_cwd = Path.cwd()
    os.chdir(output_dir)
    try:
        create_edge_router_configs_main(excel_file)
        create_sw_configs_main(excel_file)
        generate_ansible_inventory(output_dir, output_dir / "../../ansible_folder")
    finally:
        os.chdir(org_cwd)


if __name__ == "__main__":
    main()
