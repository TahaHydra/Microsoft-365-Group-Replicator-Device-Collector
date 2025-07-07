# Microsoft 365 Group Replicator & Device Collector

This PowerShell script automates the creation of a new Microsoft 365 security group based on an existing group, **filtering only human users (licensed)** and then **adding all their devices** to a second group.

## 🔧 Features

- Supports **both Microsoft 365 groups and mail-enabled security groups** (with Exchange fallback).
- Filters only users who hold one of the following licenses:
  - `STANDARDPACK` (Microsoft 365 E1)
  - `ENTERPRISEPACK` (Microsoft 365 E3)
  - `EMS` (Enterprise Mobility + Security E3)
- Creates:
  - a new user group with the filtered users.
  - a new device group containing all devices owned or registered to those users.
- Full logging to a timestamped log file.
- Automatically falls back to alternate lookup methods if Graph API calls return nothing.

## 📁 CSV Format (Input)

Your input file must be named `groups.csv` and structured like this:

```csv
nameofgroup,name2
Exco SAS,GRP_EXCO_SAS_HUMAIN
