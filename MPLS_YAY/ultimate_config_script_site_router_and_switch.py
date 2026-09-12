import os
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
NETWORK_DEV_SCRIPTS = BASE_DIR / "networkDevScripts"
if str(NETWORK_DEV_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(NETWORK_DEV_SCRIPTS))

from site_EDGE_ROUTER_script import create_edge_router_configs_main
from site_SWITCH_script import create_sw_configs_main


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Bruk: python ultimate_config_script_site_router_and_switch.py <excel-fil>")

    excel_file = Path(sys.argv[1]).resolve()
    if not excel_file.exists():
        raise FileNotFoundError(f"Fant ikke Excel-filen: {excel_file}")

    output_dir = BASE_DIR / "networkConfigs"
    output_dir.mkdir(exist_ok=True)

    org_cwd = Path.cwd()
    os.chdir(output_dir)
    try:
        create_edge_router_configs_main(excel_file)
        create_sw_configs_main(excel_file)
    finally:
        os.chdir(org_cwd)


if __name__ == "__main__":
    main()
