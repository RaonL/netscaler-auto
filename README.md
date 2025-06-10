# NetScaler Automation Scripts

This repository contains scripts for monitoring and configuring Citrix NetScaler devices. The original implementation used a Node.js script to check NetScaler status via SSH. A new Python script `netscaler_config.py` is provided to configure NetScaler using the Nitro REST API.

## Usage

- `index.js` – Node.js monitoring tool that records device status to an Excel file.
- `netscaler_config.py` – Python script that mirrors the PowerShell configuration example and sends REST API calls to NetScaler.

Each script requires valid credentials and network access to the target device.
For `netscaler_config.py`, sensitive values should be supplied via environment variables:

```
export NEW_PASSWORD=your_nsroot_password
export LDAP_SVC_PASSWORD=your_ldap_service_password
export CTXGW_VIP=your_gateway_ip
```

Run the script using `python3 netscaler_config.py` after setting these variables.
