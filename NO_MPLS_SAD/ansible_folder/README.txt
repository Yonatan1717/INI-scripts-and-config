SAFE CHANGE BUNDLE
==================

Forventet plassering:
  /cisco/ansible/

Filer:
  safe_change.sh         - Orkestrerer pre-change -> deploy -> optional apply -> verify -> rollback
  pre_change.yml         - Tar rollback-punkt før endring
  prepare_restore.yml    - Finner nyeste pre_change og SCP-er den til flash
  restore.yml            - configure replace til pre_change + write memory
  vars.yml               - Felles backup_root
  daily_back_up.yml      - Vanlig historisk dagsbackup
  ansible.cfg            - Bruker inventory.ini automatisk
  changes/loop99/        - Ufarlig eksempel-change med deploy + verify

inventory.ini:
  Legges i samme mappe, men er IKKE inkludert fordi generatoren din lager denne automatisk.

Cisco-forutsetning:
  ip scp server enable

Kjør:
  chmod +x safe_change.sh

  ./safe_change.sh R1-SITE1 loop99

Change-struktur:
  changes/<change_name>/
    deploy.yml            (påkrevd)
    verify.yml            (påkrevd)
    apply.yml             (valgfri)

Hvis apply.yml finnes:
  safe_change.sh kjører den i en egen ansible-playbook-prosess etter deploy.yml.
  Dette er spesielt nyttig for gammel CML IOS der SCP + ny CLI-kanal i samme
  prosess tidligere ga "Failed to open_session: [-1]".

Backup-struktur:
  /cisco/backup/<hostname>/daily/
  /cisco/backup/<hostname>/pre_change/

Viktig:
  Ved feil i deploy eller verify brukes siste fil i pre_change/ til rollback.
  safe_change.sh avslutter med exit code 1 etter en vellykket rollback, slik at
  automasjon fortsatt kan se at selve endringen feilet selv om rollback lyktes.
