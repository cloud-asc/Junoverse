# Junoverse

<p align="center">
  <img src="/Pictures/Logo.png" width="300">
</p>

A simple Python parser for Junos configuration files that extracts subnet information.
## Overview

Junoverse is a Python-based network segmentation discovery and visualization tool designed for parsing JUNOS configuration files and exporting:

Version: `1.0`

- 📊 Excel subnet reports
- 🗺️ Interactive HTML network segmentation maps
- 🖨️ Categorized network objects (Servers, Printers, WiFi, Voice, DMZ, etc.)
- 🌐 Network topology-style visualization

Junoverse automatically parses JUNOS configuration files and extracts:

- Subnet names
- Interface IPs
- VLANs
- DHCP pools
- Network segments
- Interface units
- Hostnames

## HTML Network Mapping

Junoverse generates categorized HTML segmentation maps with visual icons for:

| Type | Icon |
|---|---|
| Users | 💻 |
| Servers | 🖥️ |
| Printers | 🖨️ |
| WiFi | 📡 |
| Voice | ☎️ |
| Cameras | 📷 |
| IoT | 🔌 |
| Management | 🛠️ |
| DMZ | 🧱 |
| Generic Network | 🌐 |

---

## Supported JUNOS Configurations

Junoverse supports:

- `interfaces { ... }`
- `set interfaces ...`
- `access address-assignment`
- DHCP pools
- VLAN interfaces
- IRB interfaces
- Loopbacks
- VRRP interfaces
- IPv4 / IPv6 addressing

---

## Installation

```bash
git clone https://github.com/nickvourd/Junoverse.git
cd Junoverse
pip3 install -r requirements.txt
```

---

## Requirements

```
pandas>=2.0.0
openpyxl>=3.1.0
```

--- 

## Usage

- Parse a Folder:

```
python3 junoverse.py -i configs -o subnets.xlsx --html network_map.html
```

- Parse a Single File:

```
python3 junoverse.py -i router.conf -o output.xlsx --html map.html
```

---

## Credits

Created with :heart: by [@nickvourd](https://x.com/nickvourd/).
