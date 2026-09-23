# Rsyslog-server på Ubuntu for Cisco

Ubuntu brukes som sentral syslog-server.

## Labverdier
- Syslog-server: `10.1.10.10`
- UDP/514
- Cisco sender fra MGMT source-interface/VRF der plattformen støtter det

## 1. Installer
```bash
sudo apt update
sudo apt install -y rsyslog
sudo systemctl status rsyslog
```

## 2. Aktiver UDP/514
```bash
sudo nano /etc/rsyslog.d/10-cisco-remote.conf
```
Sett inn:
```text
module(load="imudp")
input(type="imudp" port="514")

template(
    name="CiscoPerHost"
    type="string"
    string="/var/log/remote-cisco/%FROMHOST-IP%/syslog.log"
)

if ($fromhost-ip != "127.0.0.1") then {
    action(
        type="omfile"
        dynaFile="CiscoPerHost"
        createDirs="on"
    )
    stop
}
```

## 3. Valider og restart
```bash
sudo rsyslogd -N1
sudo systemctl restart rsyslog
sudo systemctl enable rsyslog
```

## 4. Sjekk UDP/514
```bash
sudo ss -lunp | grep ':514'
```

## 5. Cisco-switch, referanse
```cisco
service timestamps log datetime msec show-timezone
logging host 10.1.10.10 transport udp port 514
logging trap informational
logging buffered 16384 informational
logging source-interface Vlan10
```

## 6. Cisco-router med MGMT VRF, referanse
```cisco
service timestamps log datetime msec show-timezone
logging host 10.1.10.10 vrf MGMT transport udp port 514
logging trap informational
logging buffered 16384 informational
logging source-interface Loopback10 vrf MGMT
```
Eksakt VRF/source-interface-syntaks kan variere på eldre IOS.

## 7. Se logger
```bash
sudo find /var/log/remote-cisco -maxdepth 2 -type f -print
sudo find /var/log/remote-cisco -name syslog.log -exec tail -F {} +
```

## 8. Se pakker
```bash
sudo tcpdump -ni any udp port 514
```
På Cisco:
```cisco
show logging
```

## 9. Feilsøking
Hvis pakker sees i tcpdump, men filer ikke opprettes:
```bash
sudo rsyslogd -N1
sudo journalctl -u rsyslog -n 100 --no-pager
```
Hvis ingen pakker kommer: sjekk routing, source-interface/VRF, Ubuntu-DGW, ACL/FLOW_POLICY og UFW.

Syslog er separat fra ERSPAN: syslog sender hendelseslogger, mens ERSPAN speiler nettverkstrafikk til IDS/sensor.
