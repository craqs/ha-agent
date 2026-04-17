# HA Agent

A lightweight Windows tray app that exposes your PC's state to **Home Assistant** via MQTT auto-discovery.

## What it reports

| Entity | Type | Description |
|--------|------|-------------|
| Camera Active | `binary_sensor` | `ON` while any app is using the webcam |
| Microphone Active | `binary_sensor` | `ON` while any app is using the mic |
| `<hostname>` PC | `device_tracker` | `home` while the PC is on, `not_home` after shutdown |
| `<hostname>` Notify | `notify` | Send desktop toast notifications from HA automations |

All entities group under a single HA device named after your PC's hostname.

## Installation

1. Download **ha-agent.exe** from the [latest release](../../releases/latest).
2. Run it — a setup wizard opens on first launch.
3. Enter your MQTT broker address (or click **Auto-discover** to find it via mDNS), port, and credentials.
4. Tick **Start automatically with Windows** if you want it to run at login.
5. Click **Save & Connect** — the agent appears in the system tray and entities show up in HA within seconds.

## Requirements

- Windows 10 / 11
- Home Assistant with the **MQTT integration** enabled (broker must be reachable from the PC)

## Tray menu

| Item | Action |
|------|--------|
| Status | Shows current broker connection and sensor states |
| Settings… | Re-opens the setup wizard |
| Exit | Disconnects and removes the tray icon |

## Uninstall autostart

```
ha-agent.exe --uninstall
```

## Building from source

```
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name ha-agent ha_agent.py
```

Or push a `v*` tag to GitHub — Actions builds the `.exe` automatically on a Windows runner and attaches it to a release.

## Sending notifications from HA

After the agent connects, a `notify.<hostname>_notify` service appears in Home Assistant.
Use it in automations or the Developer Tools → Services panel:

```yaml
action:
  - service: notify.desktop_abc_notify
    data:
      title: "Motion detected"
      message: "Front door camera triggered"
```

The notification appears as a balloon from the system tray icon.

> **Note:** MQTT notify entity discovery requires Home Assistant 2024.x or later.
> On older versions, add this to `configuration.yaml` manually:
> ```yaml
> notify:
>   - platform: mqtt
>     name: "Desktop ABC Notify"
>     command_topic: "ha_agent/DESKTOP-ABC/notify"
> ```

## How camera/mic detection works

The agent polls the Windows `CapabilityAccessManager` registry hive every 5 seconds.
When an app opens the camera or microphone, Windows sets `LastUsedTimeStop = 0` for that app's entry.
The agent reads this flag without requiring admin rights.
