# BTC Tools ini patcher (ipRangeGroups)

Small GUI utility to patch `ipRangeGroups` in `BTC_Tools.ini`
using a list of **explicit IP addresses**.

Designed as a lightweight helper for BTC Tools users
who manage large ASIC fleets and want reproducible scan lists.

## Features
- GUI (no CLI needed)
- Supports `.txt`, `.csv`, `.xlsx`
- Extracts and normalizes IPv4 addresses
- Generates `ipRangeGroups` using **explicit IP lists**
- Optional splitting into multiple groups (`#LAN1`, `#LAN2`, ...)
- Automatic backup of original ini
- Single-file `.exe` build (PyInstaller)

## Example output

```ini
ipRangeGroups="#LAN:10.1.1.10,10.1.1.11,10.1.1.12"
Or (large lists):
ipRangeGroups="#LAN1:10.1.1.10,10.1.1.11;#LAN2:10.1.1.12,10.1.1.13"
```
Build (Windows)
py -m pip install pyinstaller openpyxl

py -m PyInstaller --onefile --windowed --name BTCToolsIniPatcher btctools_ini_patcher_gui.py
