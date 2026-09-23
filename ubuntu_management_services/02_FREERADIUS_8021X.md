# FreeRADIUS + 802.1X på Ubuntu

Ubuntu brukes som RADIUS-server for Cisco-switcher og IEEE 802.1X.

## Labverdier
- RADIUS-server: `10.1.10.10`
- Shared secret: `RADIUS-KEY`
- UDP/1812 auth, UDP/1813 accounting
- Eksempel SW1 Site 1: `10.1.10.2`
- Eksempel SW1 Site 2: `10.2.10.2`

Viktig: RADIUS-klienten sett fra FreeRADIUS er switchen, ikke sluttbruker-PC-en.

## 1. Installer
```bash
sudo apt update
sudo apt install -y freeradius freeradius-utils
sudo systemctl status freeradius
```

## 2. Legg inn switchene som clients
Rediger:
```bash
sudo nano /etc/freeradius/3.0/clients.conf
```
Eksempel:
```text
client SW1-SITE-1 {
    ipaddr = 10.1.10.2
    secret = RADIUS-KEY
    shortname = SW1-SITE-1
}

client SW1-SITE-2 {
    ipaddr = 10.2.10.2
    secret = RADIUS-KEY
    shortname = SW1-SITE-2
}
```
Legg til én blokk per switch. IP-en må matche switchens faktiske RADIUS source-IP, i dette designet MGMT SVI/Vlan10.

## 3. Lag labbruker
```bash
sudo nano /etc/freeradius/3.0/mods-config/files/authorize
```
Legg til:
```text
testuser Cleartext-Password := "testpass"
```

## 4. Valider og debug
```bash
sudo freeradius -XC
sudo systemctl stop freeradius
sudo freeradius -X
```
Suksess viser typisk `Access-Request` og `Access-Accept`. Feil viser `Access-Reject`.

Etter debug:
```bash
sudo systemctl start freeradius
sudo systemctl enable freeradius
```

## 5. Cisco-side, referanse
```cisco
radius server CLIENT-RADIUS
 address ipv4 10.1.10.10 auth-port 1812 acct-port 1813
 key RADIUS-KEY

aaa group server radius DOT1X-RADIUS
 server name CLIENT-RADIUS

ip radius source-interface Vlan10
aaa authentication dot1x default group DOT1X-RADIUS
aaa authorization network default group DOT1X-RADIUS
aaa accounting dot1x default start-stop group DOT1X-RADIUS
dot1x system-auth-control
```
På accessport:
```cisco
interface GigabitEthernet0/1
 switchport mode access
 switchport access vlan 30
 authentication port-control auto
 dot1x pae authenticator
```
På eldre IOS kan `dot1x port-control auto` brukes.

## 6. Ubuntu-klient for labtest
```bash
sudo apt install -y wpasupplicant
sudo nano /etc/wpa_supplicant/wired.conf
```
```text
ctrl_interface=/run/wpa_supplicant
ap_scan=0
network={
    key_mgmt=IEEE8021X
    eap=PEAP
    identity="testuser"
    password="testpass"
    phase2="auth=MSCHAPV2"
    eapol_flags=0
}
```
Start på riktig interface, eksempel `ens3`:
```bash
sudo wpa_supplicant -D wired -i ens3 -c /etc/wpa_supplicant/wired.conf -dd
```
Suksess: `EAPOL authentication completed - result=SUCCESS`.

## 7. Verifiser på switch
```cisco
show authentication sessions
```
eller:
```cisco
show dot1x all
show dot1x interface GigabitEthernet0/1 details
```
Se etter `Authorized`.

## 8. UFW
Hvis UFW brukes, tillat UDP/1812 og 1813 fra hver switch-MGMT-IP.

## 9. Produksjonsnotat
Manuell `wpa_supplicant` er bare lab. I drift bør 802.1X-profil distribueres automatisk, gjerne EAP-TLS med maskinsertifikater.
