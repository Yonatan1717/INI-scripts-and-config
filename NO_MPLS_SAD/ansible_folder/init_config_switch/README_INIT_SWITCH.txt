INIT / CONFIG SWITCH
====================

Formål:
  Lage ett midlertidig L2-stagingnett der Ansible-control-noden kan nå
  alle Cisco-enhetene etter at deres *_SSH_SCP.cfg er lagt inn via console.

Fysisk oppkobling:
  Port 1         -> Ansible control node (ACCESS VLAN 10)
  Neste porter   -> Site-routere (TRUNK, tagged VLAN 10)
  Resterende     -> Site-switcher (ACCESS VLAN 10)

Hvorfor router-portene er trunk:
  Router-bootstrapen legger MGMT-adressen på et dot1Q-subinterface.
  VLAN 10 må derfor komme tagged inn til routeren.

VIKTIG om control node:
  Site 1, Site 2, Site 3 osv. bruker ulike MGMT-IP-subnett selv om alle
  bruker VLAN 10. INIT-switchen ruter ikke mellom dem. Control-noden må
  derfor ha én midlertidig IP-adresse i hvert MGMT-subnett på samme NIC.
  CONTROL_NODE_TEMP_IPS.sh inneholder forslag til slike adresser.

Arbeidsflyt:
  1. Velg riktig INIT_CONFIG_SWITCH_*.cfg og konfigurer staging-switchen.
  2. Koble Ansible-serveren til port 1.
  3. Legg midlertidige site-MGMT-adresser på Ansible-NIC-en.
  4. Koble routere til trunkportene og switcher til accessportene.
  5. Lim inn riktig *_SSH_SCP.cfg via console på hver target-enhet.
  6. Verifiser SSH fra Ansible-serveren.
  7. Kjør deploy_generated.yml.
  8. Verifiser og kjør write memory.
  9. Flytt ferdig konfigurerte enheter til endelig topologi.

Dette er et midlertidig stagingnett og skal ikke stå parallelt koblet mot
produksjonsnettet under førstegangsoppsettet.
