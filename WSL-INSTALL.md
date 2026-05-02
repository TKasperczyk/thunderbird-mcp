# WSL Installation Guide for Thunderbird MCP

This guide explains how to connect the Thunderbird MCP server running on Windows to a WSL2 environment.

## Prerequisites

- Thunderbird running on Windows with the MCP extension installed
- WSL2 installed and running
- The extension's **"Listen on all interfaces"** option enabled (Add-ons → Thunderbird MCP → Preferences)

## Problem

WSL2 uses NAT networking by default. The Thunderbird MCP server listens on `0.0.0.0:8765` on Windows, but WSL2 cannot reach it directly because:
1. Windows Firewall blocks inbound connections from the WSL2 subnet
2. WSL2's NAT layer doesn't forward TCP traffic to the host's virtual interface

## Solution: Port Proxy + Firewall Rule

### Step 1: Allow Inbound Firewall Traffic

Open PowerShell as **Administrator** and run:

```powershell
New-NetFirewallRule -DisplayName "Thunderbird MCP" -Direction Inbound -Protocol TCP -LocalPort 8765 -Action Allow -Profile Any
```

### Step 2: Find Your WSL2 Gateway IP

From WSL2, run:

```bash
WSL_GATEWAY=$(grep nameserver /etc/resolv.conf | awk '{print $2}')
echo $WSL_GATEWAY
```

This gives you the IP of your WSL2 NAT gateway (e.g., `172.18.64.1`).

### Step 3: Set Up Port Forwarding

In PowerShell as **Administrator**, add a port proxy rule using your gateway IP:

```powershell
netsh interface portproxy add v4tov4 listenport=8765 listenaddress=<WSL_GATEWAY_IP> connectport=8765 connectaddress=127.0.0.1
```

This forwards traffic from the WSL2 gateway IP to localhost where Thunderbird is listening.

### Step 4: Verify the Setup

From WSL2, test the connection:

```bash
WSL_GATEWAY=$(grep nameserver /etc/resolv.conf | awk '{print $2}')
curl -s --http1.0 http://$WSL_GATEWAY:8765/ \
  -H "Authorization: Bearer $(cat /mnt/c/Users/<YOUR_USER>/AppData/Local/Temp/thunderbird-mcp/connection.json | jq -r .token)" \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list"}'
```

You should see a JSON response listing available tools.

## Using the WSL Bridge

The `mcp-bridge.wsl.cjs` script is designed for WSL2 environments. It automatically detects the Windows gateway IP.

### Quick Start

```bash
# Set the connection file path (adjust username)
export THUNDERBIRD_MCP_CONNECTION_FILE="/mnt/c/Users/<YOUR_USER>/AppData/Local/Temp/thunderbird-mcp/connection.json"

# Run the bridge
node mcp-bridge.wsl.cjs
```

### With Bun

```bash
THUNDERBIRD_MCP_CONNECTION_FILE="/mnt/c/Users/<YOUR_USER>/AppData/Local/Temp/thunderbird-mcp/connection.json" bun mcp-bridge.wsl.cjs
```

### MCP Client Configuration

In your MCP client config (e.g., Claude Desktop, Cursor), point to the bridge:

```json
{
  "mcpServers": {
    "thunderbird": {
      "command": "node",
      "args": ["/path/to/mcp-bridge.wsl.cjs"],
      "env": {
        "THUNDERBIRD_MCP_CONNECTION_FILE": "/mnt/c/Users/<YOUR_USER>/AppData/Local/Temp/thunderbird-mcp/connection.json"
      }
    }
  }
}
```

## Troubleshooting

### Connection timed out
- Verify the firewall rule exists: `Get-NetFirewallRule -DisplayName "Thunderbird MCP"`
- Verify the portproxy rule: `netsh interface portproxy show all`
- Check that "Listen on all interfaces" is enabled in the extension

### Bad request (400)
- The bridge uses HTTP/1.0 internally to work through `netsh portproxy`. If you're testing with curl, use `--http1.0` flag

### Authentication failed (403)
- The connection file may be stale. Thunderbird regenerates the token on restart
- Re-read the connection file or restart Thunderbird

### Portproxy persists across reboots
The `netsh interface portproxy` rule is persistent by default and survives reboots. To remove it:
```powershell
netsh interface portproxy delete v4tov4 listenport=8765 listenaddress=<WSL_GATEWAY_IP>
```

## Alternative: Mirrored Networking Mode

If you don't use Docker Desktop, you can enable WSL2 mirrored networking which eliminates the NAT layer entirely:

```ini
# /etc/wsl.conf
[experimental]
networkingMode=mirrored
```

Then restart WSL2 (`wsl --shutdown` from PowerShell). With mirrored mode, `localhost:8765` works directly and no portproxy is needed.

**Warning:** Mirrored mode is incompatible with Docker Desktop (versions 4.39+). See [Docker issue #14691](https://github.com/docker/for-win/issues/14691).
