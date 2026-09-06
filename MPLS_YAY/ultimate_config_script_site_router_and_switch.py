import os
import sys
import pathlib

sys.path.append("networkDevScripts")

from site_SWITCH_script import create_sw_configs_main
from site_EDGE_ROUTER_script import create_edge_router_configs_main


org_cwd = os.getcwd()

excel_files = sys.argv[1]
abs_path = pathlib.Path(excel_files).resolve()

if not os.path.exists("networkConfigs"):
    os.mkdir("networkConfigs")

os.chdir("networkConfigs")

try:
    create_edge_router_configs_main(abs_path)
    
    create_sw_configs_main(abs_path)
finally:
    os.chdir(org_cwd)

