import os
import sys
import json
from numpy import rint
import pandas as pd
from openpyxl import load_workbook
import ipaddress

SSH_DOMAIN = "lab.local"
have_asked_rsyslog_server = False
have_asked_tacacs_server = False
isp_added = False
isp_added2 = False

def read_sheet(filename, sheet):
    df = pd.read_excel(
        filename,
        sheet_name=sheet,
        header=None
    )

    md_start = 0
    md_end = df.iloc[md_start:].isna().all(axis=1).idxmax()
    md = df.iloc[md_start:md_end, 0:8]
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
    
    md_swi_start = tunnel_data_end + 1
    blank = df.iloc[md_swi_start:].isna().all(axis=1)
    md_swi_end = blank.idxmax()
    md_swi = df.iloc[md_swi_start:md_swi_end, 0:6]
    md_swi.columns = md_swi.iloc[0]
    md_swi = md_swi[1:].reset_index(drop=True)

    
    swi_data_start = md_swi_end + 1

    swi_data = df.iloc[swi_data_start:, 0:10]

    swi_data.columns = swi_data.iloc[0]
    swi_data = swi_data[1:].reset_index(drop=True)
    
    
    return {
        "md": md_swi,
        "swi_data": swi_data,
        "md_top": md,
    }

    # Removed redundant return statement


def create_tacacs_config(md_top):
    my_data = {}
    my_data["config"] = {}
    my_data["network_info"] = {}

    global have_asked_tacacs_server
    global tacacs_server
    if not have_asked_tacacs_server:
        tacacs_server = input("IP-adressen til TACACS-serveren: ")
        have_asked_tacacs_server = True

    tacacs_key = md_top.iloc[0].get("tacacs_key", "")

    if not tacacs_server or not tacacs_key:
        print("tacas feila")
        exit(1)

    my_data["config"]["aaa new-model"] = []
    my_data["config"][f"aaa group server tacacs+ TACACS-GROUP"] = [
        f"server-private {tacacs_server} key {tacacs_key}",
        f"ip tacacs source-interface Vlan10",
        "exit"
    ]
    my_data["config"][f"aaa authentication login default group TACACS-GROUP local"] = []
    my_data["config"][f"aaa authorization exec default group TACACS-GROUP local"] = []

    return my_data


def create_rsyslog_config():
    my_data = {}
    my_data["config"] = {}
    my_data["network_info"] = {}

    global have_asked_rsyslog_server
    global rsyslog_server
    
    if not have_asked_rsyslog_server:
        rsyslog_server = input("IP-adressen til Rsyslog-serveren: ")
        have_asked_rsyslog_server = True
    if not have_asked_rsyslog_server:
        rsyslog_server = input("IP-adressen til Rsyslog-serveren: ")

    if not rsyslog_server:
        print("rsyslog server not found")
        exit(1)

    my_data["config"][f"service timestamps log datetime msec show-timezone"] = []
    my_data["config"][f"logging host {rsyslog_server} transport udp port 514"] = []
    my_data["config"][f"logging trap informational"] = []
    my_data["config"][f"logging source-interface Vlan10"] = []

    return my_data


def enable_ssh(md, domain=SSH_DOMAIN):
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
    my_data["config"]["crypto key generate rsa general-keys modulus 4096"] = []
    my_data["config"]["ip ssh version 2"] = []
    my_data["config"][f"line vty {' '.join(x.strip(' ') for x in vty_lines.split('-'))}"] = [
        "login authentication default",
        "transport input ssh",
        "exit",
    ]

    return my_data


def global_config(md,md_top, swi_data):
    info = {}
    info["config"] = {}
    info["network_info"] = {}

    site = md.iloc[0]["site"]
    secret = md.iloc[0].get("secret", "")

    for idx, row in swi_data.iterrows():
        sw_id = row["SW"]
        mgmt_vlan = row["MGMT Vlan"]
        mgmt_ip = row["MGMT ip"]
        gatway = row["gateway"]
        mask = row["mask"]
        intf_prefix = row["intf_prefix"]

        if f"SW{sw_id}-SITE-{site}" not in info["config"]:
            info["config"][f"SW{sw_id}-SITE-{site}"] = {}

        info["config"][f"SW{sw_id}-SITE-{site}"][f"hostname SW{sw_id}-SITE-{site}"] = []
        info["config"][f"SW{sw_id}-SITE-{site}"][f"enable secret 9 {secret}"] = []

        tacacs_config = create_tacacs_config(md_top) 
        info["config"][f"SW{sw_id}-SITE-{site}"].update(tacacs_config["config"])
        ssh_config = enable_ssh(md)
        info["config"][f"SW{sw_id}-SITE-{site}"].update(ssh_config["config"])
        rsyslog_config = create_rsyslog_config()
        info["config"][f"SW{sw_id}-SITE-{site}"].update(rsyslog_config["config"])


        info["config"][f"SW{sw_id}-SITE-{site}"][f"vlan {mgmt_vlan}"] = [
            f"name MGMT_VLAN_{mgmt_vlan}",
            "exit"
        ]
        info["config"][f"SW{sw_id}-SITE-{site}"][f"interface vlan {mgmt_vlan}"] = [
            f"ip address {mgmt_ip} {mask}",
            "no shutdown",
            "exit"
        ]

        info["config"][f"SW{sw_id}-SITE-{site}"][f"interface {intf_prefix}0"] = [
            f"description Management interface for VLAN {mgmt_vlan}",
            "ip arp inspection trust",
            "switchport mode access",
            f"switchport access vlan {mgmt_vlan}",
            "switchport port-security",
            "switchport port-security maximum 2",
            "switchport port-security violation restrict",
            "spanning-tree bpduguard enable",
            "spanning-tree portfast",
            "no shutdown",
            "exit"
        ]

        info["config"][f"SW{sw_id}-SITE-{site}"][f"ip default-gateway {gatway}"] = []
        info["config"][f"SW{sw_id}-SITE-{site}"][f"ntp server {gatway}"] = []


    return info


def config_vlan(swi_data, site, md):
    info = {} 
    info["config"] = {}
    info["network_info"] = {}

    for idx, row in swi_data.iterrows():
        sw_id = row["SW"]
        vlan_count = str(row["vlan-antall"])
        if "-" in vlan_count:
            vlan_info = [tuple(map(int, x.split("."))) for x in vlan_count.split("-")]
        else:
            vlan_info = [tuple(map(int, vlan_count.split(".")))]
            
        # print(vlan_info)
        # exit()

        intf_prefix = row["intf_prefix"]

        global isp_added2

        if not isp_added2:
            isp_vlan = md.iloc[0]["ISP-VLAN"]
            isp_added2 = True
        else:
            isp_vlan = "hello"

        if f"SW{sw_id}-SITE-{site}" not in info["config"]:
            info["config"][f"SW{sw_id}-SITE-{site}"] = {}

        info["config"][f"SW{sw_id}-SITE-{site}"][f"vlan 999"] = [
            f"name NATIVE_UBRUKT",
            "exit"
        ]

        made = 0
        for vlan, antall in vlan_info:
            if antall <= 0:
                continue
            info["config"][f"SW{sw_id}-SITE-{site}"][f"vlan {vlan}"] = [
                f"name VLAN_{vlan}",
                "exit"
            ]

            rng = f"{made + 1}-{made + antall}" if antall > 1 else f"{made + 1}"
            range_or_not = "range " if antall > 1 else "" 

            info["config"][f"SW{sw_id}-SITE-{site}"][f"interface {range_or_not}{intf_prefix}{rng}"] = [
                f"description access port for VLAN {vlan}",
                "ip dhcp snooping trust" if vlan == isp_vlan else "!",
                "ip arp inspection trust" if vlan == isp_vlan else "!",
                "switchport mode access",
                f"switchport access vlan {vlan}",
                "switchport port-security",
                "switchport port-security maximum 2",
                "switchport port-security violation restrict",
                "spanning-tree bpduguard enable",
                "spanning-tree portfast",
                "no shutdown",
                "exit"
            ]

            made += antall


    return info


def config_trunk_and_dchp_snooping(swi_data, site, md):
    info = {} 
    info["config"] = {}
    info["network_info"] = {}

    for idx, row in swi_data.iterrows():
        sw_id = row["SW"]
        mgmg_vlan = int(row["MGMT Vlan"])
        vlan_count = str(row["vlan-antall"])
        if "-" in vlan_count:
            vlan_info = [tuple(map(int, x.split("."))) for x in vlan_count.split("-")]

            vlans = [vlan for vlan, antall in vlan_info]
            tot_antall_port = sum([int(antall) for vlan, antall in vlan_info])
        else:
            vlans = [int(vlan_count.split(".")[0])]
            antall = int(vlan_count.split(".")[1])
            tot_antall_port = antall

        

       
        if mgmg_vlan not in vlans:
            vlans.insert(0, mgmg_vlan)
            
        intf_prefix = row["intf_prefix"]
        num_ports = row["num_ports"]
        if type(num_ports) is not int:
            num_ports = int(num_ports)

        if f"SW{sw_id}-SITE-{site}" not in info["config"]:
            info["config"][f"SW{sw_id}-SITE-{site}"] = {}

        to_lan = num_ports - 2
        to_core = num_ports - 1
        
        vlans.append(999)
        
        global isp_added
        
        if not isp_added:
            try:
                isp_vlan = md.iloc[0]["ISP-VLAN"]
                if type(isp_vlan) is not int:
                    isp_vlan = int(isp_vlan)
                vlans_for_dhcp_snooping = [vlan for vlan in vlans if vlan != isp_vlan]
            except Exception as e:
                vlans_for_dhcp_snooping = vlans

            isp_added = True
        else:
            vlans_for_dhcp_snooping = vlans


        info["config"][f"SW{sw_id}-SITE-{site}"][f"ip dhcp snooping"] = []
        info["config"][f"SW{sw_id}-SITE-{site}"][f"ip dhcp snooping vlan {','.join(map(str, vlans_for_dhcp_snooping))}"] = []
        info["config"][f"SW{sw_id}-SITE-{site}"][f"no ip dhcp snooping information option"] = []
        info["config"][f"SW{sw_id}-SITE-{site}"][f"ip arp inspection vlan {','.join(map(str, vlans_for_dhcp_snooping))}"] = []

        info["config"][f"SW{sw_id}-SITE-{site}"][f"interface {intf_prefix}{to_core}"] = [
            f"description uplink trunk port for VLAN {','.join(map(str, vlans))}",
            "switchport trunk encapsulation dot1q",
            "switchport trunk native vlan 999",
            "switchport mode trunk",
            f"switchport trunk allowed vlan {','.join(map(str, vlans))}",
            f"ip dhcp snooping trust",
            f"ip arp inspection trust",
            "no shutdown",
            "exit"
        ]

        num_down_ports = row["num_downlink"]
        # print(type(num_down_ports))
        # exit()
        for i in range(num_down_ports):
            info["config"][f"SW{sw_id}-SITE-{site}"][f"interface {intf_prefix}{to_lan - i}"] = [
                f"description downlink trunk port for VLAN {','.join(map(str, vlans))}. OM ikke brukt skal det brukes shutdown på porten",
                "switchport trunk encapsulation dot1q",
                "switchport trunk native vlan 999",
                "switchport mode trunk",
                f"switchport trunk allowed vlan {','.join(map(str, vlans))}",
                f"ip dhcp snooping trust",
                f"ip arp inspection trust",
                "no shutdown",
                "exit"
            ]
    
        ports_left = num_ports - tot_antall_port - 2 - num_down_ports
        if ports_left > 0:
            start_int = tot_antall_port + 1
            end_int = num_ports - 2 - num_down_ports
            range_or_not = "range " if start_int != end_int else ""

            rng = f"{start_int}-{end_int}" if start_int != end_int else f"{start_int}"

            info["config"][f"SW{sw_id}-SITE-{site}"][f"interface {range_or_not}{intf_prefix}{rng}"] = [
                f"description ubrukt port access ports",
                "shutdown",
                "exit"
            ]


    return info


def update_site_config(data,swi_data, sn, conf):
    for sw_id in swi_data["SW"].unique():
        if f"SW{sw_id}-SITE-{sn}" not in data[f"site {sn}"]["config"]:
            data[f"site {sn}"]["config"][f"SW{sw_id}-SITE-{sn}"] = {}

        data[f"site {sn}"]["config"][f"SW{sw_id}-SITE-{sn}"].update(conf["config"][f"SW{sw_id}-SITE-{sn}"])

    return data


def fetch_site_data(config_file):
    try:
        data = {}

        with open(config_file, "r") as f:
            try:
                data = json.load(f)
            except json.JSONDecodeError:
                data = {}

    except FileNotFoundError:
        data = {}

    return data


def create_site_sw_config(file, sheet, config_file):
    sheet_data = read_sheet(file, sheet)
    data = fetch_site_data(config_file)

    md = sheet_data["md"]
    md_top = sheet_data["md_top"]
    swi_data = sheet_data["swi_data"]

    sn = md.iloc[0]["site"]
    data[f"site {sn}"] = {}
    data[f"site {sn}"]["config"] = {}

    vlan_conf = config_vlan(swi_data, sn, md)
    data = update_site_config(data, swi_data, sn, vlan_conf)

    global_conf = global_config(md, md_top, swi_data)
    data = update_site_config(data, swi_data, sn, global_conf)

    trunk_and_dchp_snooping_conf = config_trunk_and_dchp_snooping(swi_data, sn, md)
    data = update_site_config(data, swi_data, sn, trunk_and_dchp_snooping_conf)

    with open(config_file, "w") as f:
        f.write(json.dumps(data, indent=4))


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
    if not os.path.exists("siteSwichTextConfigs"):
        os.makedirs("siteSwichTextConfigs")

        
    for site, site_data in data.items():
        if not os.path.exists(f"siteSwichTextConfigs/site_{site}"):
            os.makedirs(f"siteSwichTextConfigs/site_{site}")

        configs = site_data["config"]

        for sw_name, config in configs.items():
            text = config_to_text(config)

            with open(
                f"siteSwichTextConfigs/site_{site}/{sw_name}.txt",
                "w",
                encoding="utf-8"
            ) as f:
                f.write("\n".join(text))
            

    print(f"Text versjon av config for svitjer i site {site}, har blitt lagret i siteSwichTextConfigs/site_{site}/{sw_name}.txt")



def create_sw_configs_main(file, config_file="site_switch_config.json"):   

    sites_sheets = load_workbook(file).sheetnames
    
    for sheet in sites_sheets:
        data = create_site_sw_config(file, sheet, config_file)
    
    create_or_update_config_files(data)
   
        
def main():
    file = sys.argv[1]
    config_file = "site_switch_config.json" if len(sys.argv) < 3 else sys.argv[2]
    create_sw_configs_main(file, config_file)


if __name__ == "__main__":
    main()