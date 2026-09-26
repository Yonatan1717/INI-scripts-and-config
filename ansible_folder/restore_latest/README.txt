restore_latest
==============

Expected location:
  /cisco/ansible/changes/restore_latest/

deploy.yml:
  Finds the newest normal .cfg backup directly under /cisco/backup/<host>/
  and copies it to flash:/ansible-restore.cfg.

apply.yml:
  Runs configure replace against the staged file, then write memory.

verify.yml:
  Compares flash:/ansible-restore.cfg with system:running-config and fails if
  Cisco reports + or - diff lines.

Why apply.yml is separate:
  The legacy CML IOS image in this lab failed when SCP and the next CLI channel
  were used in the same Ansible process. Keep deploy and apply as separate
  ansible-playbook runs.

Example:
  ansible-playbook changes/restore_latest/deploy.yml --limit R1-SITE1
  rm -rf ~/.ansible/pc/*
  ansible-playbook changes/restore_latest/apply.yml --limit R1-SITE1
  ansible-playbook changes/restore_latest/verify.yml --limit R1-SITE1

Required on IOS:
  ip scp server enable

/cisco/ansible/vars.yml:
  ---
  backup_root: "/cisco/backup"

deploy.yml uses recurse: false intentionally so pre_change/ is not selected.
