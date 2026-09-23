# SNMPv3 / snmpwalk på Ubuntu

Denne filen setter opp Ubuntu som enkel NMS-klient for SNMPv3-polling mot Cisco-routere og -switcher.

## Labverdier
- NMS/Ubuntu-IP: `10.1.10.10`
- Eksempel-switch: `10.1.10.2`
- SNMPv3-bruker: `nmsuser`
- Auth: SHA
- Auth-passord: `SNMP-AUTH-KEY`
- Privacy: AES-128
- Privacy-passord: `SNMP-PRIV-KEY`
- Polling: UDP/161

Bytt passordene før fysisk/produksjonsnær bruk.

## 1. Installer
```bash
sudo apt update
sudo apt install -y snmp
```

## 2. Cisco-side, referanse
```cisco
ip access-list standard SNMP-NMS-ONLY
 permit host 10.1.10.10
 deny any

snmp-server view NMS-READ iso included
snmp-server group NMS v3 priv read NMS-READ access SNMP-NMS-ONLY
snmp-server user nmsuser NMS v3 auth sha SNMP-AUTH-KEY priv aes 128 SNMP-PRIV-KEY
```

## 3. Test med snmpwalk
```bash
snmpwalk -v3 \
  -l authPriv \
  -u nmsuser \
  -a SHA \
  -A 'SNMP-AUTH-KEY' \
  -x AES \
  -X 'SNMP-PRIV-KEY' \
  10.1.10.2 \
  1.3.6.1.2.1.1
```

## 4. Kort test med snmpget
Hostname:
```bash
snmpget -v3 -l authPriv -u nmsuser -a SHA -A 'SNMP-AUTH-KEY' -x AES -X 'SNMP-PRIV-KEY' 10.1.10.2 1.3.6.1.2.1.1.5.0
```
Uptime:
```bash
snmpget -v3 -l authPriv -u nmsuser -a SHA -A 'SNMP-AUTH-KEY' -x AES -X 'SNMP-PRIV-KEY' 10.1.10.2 1.3.6.1.2.1.1.3.0
```

## 5. Feilsøking
Sjekk source-IP:
```bash
ip route get 10.1.10.2
```
Den bør vise `src 10.1.10.10` dersom Cisco-ACL-en bare tillater NMS-serveren.

Se pakker:
```bash
sudo tcpdump -ni any udp port 161
```

På Cisco:
```cisco
show ip access-lists SNMP-NMS-ONLY
show snmp
show snmp user
show snmp group
```

Ved timeout: sjekk source-IP, routing/default gateway, ACL, brukernavn, auth/priv-passord og SHA/AES-innstillinger.
