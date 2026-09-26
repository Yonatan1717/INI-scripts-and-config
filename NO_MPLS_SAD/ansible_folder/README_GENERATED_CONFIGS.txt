Generated Ansible files
=======================

This directory is a self-contained portable Ansible bundle.
The paths in inventory.ini are RELATIVE paths, not Windows/Linux absolute paths.
You can therefore copy the whole directory to the same or another host.

inventory.ini contains two per-device variables:
  config_file           = relative full Ansible-ready configuration
  bootstrap_config_file = relative console-paste MGMT + SSH/SCP bootstrap

Bootstrap is intended to make the device reachable by Ansible first.
It uses local VTY authentication and does not depend on TACACS/RADIUS.
After SSH works, a playbook can apply the full file with:

  cisco.ios.ios_config:
    src: '{{ inventory_dir }}/{{ config_file }}'
    backup: true
    save_when: modified

Important: ios_config src performs a merge. A command that disappears
from the generated file is not automatically negated on the device.
