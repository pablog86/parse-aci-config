# Parse ACI Config :white_check_mark:

## Description

This example script generates an Excel file with the information gathered from ACI backup in JSON format.

The file contains the following pages:
- Tenants policies
- Access Policies


## Requirements

Backup file in JSON format.

### Clone the repository

```text
git clone https://github.com/pablog86/parse-aci-config/
cd parse_aci_config

chmod 755 parse_aci_config
```

### Python environment

Create virtual environment and activate it (optional)

```text
python3 -m venv parse_aci_config
source parse_aci_config/bin/activate
Install required modules
```

Install required modules

```text
pip install -r requirements.txt
```


## Usage & examples

Just run the parse-conf.py script and select the backup file. The output will be generated in the same directory.
