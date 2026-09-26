import json
import os
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
