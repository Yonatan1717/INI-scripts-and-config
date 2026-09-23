# TACACS+ på Ubuntu

Ubuntu brukes som TACACS+ server for Cisco-routere og -switcher.

## Labverdier
- TACACS-server: `10.1.10.10`
- Shared key: `TACACS-KEY`
- TCP/49
- Routere bruker MGMT Loopback10 som source
- Switcher bruker MGMT SVI/Vlan10 som source

Viktig: serveren må ha korrekt default gateway, normalt `10.1.10.1` på Site 1, slik at svar til router-loopbacks returneres riktig.

## 1. Installer
```bash
sudo apt update
sudo apt install -y tacacs+
```
Finn filer/unit-navn dersom Ubuntu-versjonen avviker:
```bash
dpkg -L tacacs+ | grep -E 'tac_plus|conf|service'
systemctl list-unit-files | grep -i tac
```
Vanlig config-path:
```text
/etc/tacacs+/tac_plus.conf
```

## 2. Enkel config
```bash
sudo cp /etc/tacacs+/tac_plus.conf /etc/tacacs+/tac_plus.conf.bak
sudo nano /etc/tacacs+/tac_plus.conf
```
Eksempel:
```text
key = "TACACS-KEY"
accounting file = /var/log/tac_plus.acct

group = network-admins {
    default service = permit
    service = exec {
        priv-lvl = 15
    }
}

user = admin {
    member = network-admins
    login = cleartext "CHANGE-ME-TACACS-PASSWORD"
}
```
Valgfritt kan egne klienter defineres:
```text
host = 1.1.1.10 {
    key = "TACACS-KEY"
}

host = 2.2.2.10 {
    key = "TACACS-KEY"
}

host = 10.1.10.2 {
    key = "TACACS-KEY"
}
```

## 3. Filrettigheter
```bash
sudo chown root:root /etc/tacacs+/tac_plus.conf
sudo chmod 600 /etc/tacacs+/tac_plus.conf
```

## 4. Valider config
```bash
sudo tac_plus -P -C /etc/tacacs+/tac_plus.conf
```

## 5. Start/restart
Finn faktisk unit-navn:
```bash
systemctl list-unit-files | grep -i tac
```
Deretter:
```bash
sudo systemctl restart <UNIT-NAVN>
sudo systemctl enable <UNIT-NAVN>
sudo systemctl status <UNIT-NAVN>
```
Foreground-debug:
```bash
sudo tac_plus -g -d 16 -d 8 -d 64 -C /etc/tacacs+/tac_plus.conf
```

## 6. Sjekk TCP/49
```bash
sudo ss -lntp | grep ':49'
```

## 7. Cisco-router, referanse
```cisco
aaa new-model

tacacs server TACACS-SERVER
 address ipv4 10.1.10.10
 key TACACS-KEY

aaa group server tacacs+ TACACS-GROUP
 server name TACACS-SERVER
 ip vrf forwarding MGMT
 ip tacacs source-interface Loopback10

aaa authentication login default group TACACS-GROUP local
aaa authorization exec default group TACACS-GROUP local
aaa accounting exec default start-stop group TACACS-GROUP
aaa accounting commands 15 default start-stop group TACACS-GROUP
```

## 8. Cisco-switch, referanse
```cisco
aaa new-model

tacacs server TACACS-SERVER
 address ipv4 10.1.10.10
 key TACACS-KEY

aaa group server tacacs+ TACACS-GROUP
 server name TACACS-SERVER
 ip tacacs source-interface Vlan10

aaa authentication login default group TACACS-GROUP local
aaa authorization exec default group TACACS-GROUP local
aaa accounting exec default start-stop group TACACS-GROUP
aaa accounting commands 15 default start-stop group TACACS-GROUP
```

## 9. Test og feilsøking
Router:
```cisco
ping vrf MGMT 10.1.10.10 source Loopback10
show aaa servers
debug tacacs
```
Stopp debug:
```cisco
undebug all
```
Hvis vanlig ping fungerer, men Loopback10-source feiler, sjekk server-DGW, ACL/FLOW_POLICY og returroute.

Hvis UFW brukes, tillat TCP/49 fra de faktiske TACACS source-IP-ene, inkludert router-loopbacks og switch-MGMT-adresser.
