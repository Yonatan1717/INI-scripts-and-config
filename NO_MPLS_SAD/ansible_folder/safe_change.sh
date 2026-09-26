#!/usr/bin/env bash
set -u
set -o pipefail

ANSIBLE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INVENTORY="$ANSIBLE_DIR/inventory.ini"

TARGET="${1:-cisco}"
CHANGE="${2:-loop99}"

CHANGE_DIR="$ANSIBLE_DIR/changes/$CHANGE"
DEPLOY_PLAYBOOK="$CHANGE_DIR/deploy.yml"
APPLY_PLAYBOOK="$CHANGE_DIR/apply.yml"
VERIFY_PLAYBOOK="$CHANGE_DIR/verify.yml"

run_playbook() {
    local playbook="$1"
    ansible-playbook -i "$INVENTORY" "$playbook" --limit "$TARGET"
}

clear_connections() {
    rm -rf "$HOME/.ansible/pc/"* 2>/dev/null || true
}

echo "======================================"
echo "SAFE CHANGE"
echo "Target: $TARGET"
echo "Change: $CHANGE"
echo "======================================"

if [[ ! -f "$DEPLOY_PLAYBOOK" ]]; then
    echo "ERROR: Fant ikke $DEPLOY_PLAYBOOK"
    exit 2
fi

if [[ ! -f "$VERIFY_PLAYBOOK" ]]; then
    echo "ERROR: Fant ikke $VERIFY_PLAYBOOK"
    exit 3
fi

echo
echo "=== 1. PRE-CHANGE BACKUP ==="
run_playbook "$ANSIBLE_DIR/pre_change.yml"
if [[ $? -ne 0 ]]; then
    echo "PRE-CHANGE BACKUP FAILED - avbryter uten å gjøre endring."
    exit 10
fi

echo
echo "=== 2. DEPLOY / STAGE CHANGE ==="
run_playbook "$DEPLOY_PLAYBOOK"
DEPLOY_RESULT=$?

# Noen changes (f.eks. restore_latest) trenger en separat apply.yml.
# Dette er også nyttig på gammel CML IOS, siden SCP -> ny CLI-kanal
# i samme Ansible-prosess tidligere ga Failed to open_session.
if [[ $DEPLOY_RESULT -eq 0 && -f "$APPLY_PLAYBOOK" ]]; then
    echo
    echo "=== 3. APPLY CHANGE ==="
    clear_connections
    run_playbook "$APPLY_PLAYBOOK"
    DEPLOY_RESULT=$?
fi

if [[ $DEPLOY_RESULT -eq 0 ]]; then
    echo
    echo "=== 4. VERIFY CHANGE ==="
    clear_connections
    run_playbook "$VERIFY_PLAYBOOK"
    VERIFY_RESULT=$?
else
    VERIFY_RESULT=1
fi

if [[ $DEPLOY_RESULT -eq 0 && $VERIFY_RESULT -eq 0 ]]; then
    echo
    echo "======================================"
    echo "CHANGE SUCCESSFUL"
    echo "======================================"
    exit 0
fi

echo
echo "======================================"
echo "CHANGE/VERIFY FAILED - STARTING ROLLBACK"
echo "======================================"

clear_connections

echo
echo "=== 5. STAGE PRE-CHANGE BACKUP ==="
run_playbook "$ANSIBLE_DIR/prepare_restore.yml"
if [[ $? -ne 0 ]]; then
    echo "FAILED TO PREPARE ROLLBACK"
    exit 20
fi

# Gammel IOS/CML trenger separat SSH-prosess mellom SCP og configure replace.
clear_connections

echo
echo "=== 6. RESTORE PRE-CHANGE CONFIG ==="
run_playbook "$ANSIBLE_DIR/restore.yml"
if [[ $? -ne 0 ]]; then
    echo "ROLLBACK FAILED"
    exit 30
fi

echo
echo "======================================"
echo "ROLLBACK COMPLETE"
echo "======================================"
exit 1
