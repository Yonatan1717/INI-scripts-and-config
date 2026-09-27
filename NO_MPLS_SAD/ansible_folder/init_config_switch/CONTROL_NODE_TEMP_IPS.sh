#!/usr/bin/env bash
# Midlertidige adresser for Ansible-control-node under staging.
# BYTT <ANSIBLE_NIC> med riktig interface, f.eks. ens160.
# Kontroller at foreslåtte adresser er ledige før de tas i bruk.

NIC="${1-NONE}"
if [ "$NIC" = "NONE" ]; then
    echo "Usage: $0 <ANSIBLE_NIC>"
    exit 1
fi

sudo ip addr add 10.1.10.254/24 dev "$NIC"
sudo ip addr add 10.2.10.254/24 dev "$NIC"
sudo ip addr add 10.3.10.254/24 dev "$NIC"

# Verifisering:
ip -br addr show "$NIC"

# Fjern adressene etter staging dersom de ikke skal beholdes:
# sudo ip addr del 10.1.10.254/24 dev "$NIC"
# sudo ip addr del 10.2.10.254/24 dev "$NIC"
# sudo ip addr del 10.3.10.254/24 dev "$NIC"
