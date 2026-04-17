"""
HA Agent — Windows Home Assistant MQTT Agent

Exposes to Home Assistant via MQTT auto-discovery:
  • binary_sensor: camera currently in use
  • binary_sensor: microphone currently in use
  • device_tracker: PC online / offline
  • notify: receive desktop toast notifications from HA

Run with no arguments to start (shows setup wizard on first run).
Flags: --setup  open settings dialog
       --uninstall  remove Windows autostart entry and exit
"""

import argparse
import json
import logging
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from pathlib import Path

import paho.mqtt.client as mqtt
from PIL import Image, ImageDraw
import pystray
from zeroconf import Zeroconf, ServiceBrowser, ServiceListener

# winreg and tkinter are Windows stdlib
import winreg
import tkinter as tk
from tkinter import ttk, messagebox

# ── Constants ─────────────────────────────────────────────────────────────────

APP_NAME = "HA-Agent"
VERSION = "0.0.0"  # patched to the git tag by the GitHub Actions workflow before building
GITHUB_REPO = "craqs/ha-agent"
POLL_INTERVAL = 5  # seconds between sensor polls

HOSTNAME = socket.gethostname().replace(" ", "_")

AVAIL_TOPIC = f"ha_agent/{HOSTNAME}/availability"
CAMERA_TOPIC = f"ha_agent/{HOSTNAME}/camera"
MIC_TOPIC = f"ha_agent/{HOSTNAME}/microphone"
STATUS_TOPIC = f"ha_agent/{HOSTNAME}/status"
NOTIFY_TOPIC = f"ha_agent/{HOSTNAME}/notify"

# ── Config ────────────────────────────────────────────────────────────────────

CONFIG_DIR = Path(os.environ.get("APPDATA", ".")) / APP_NAME
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULT_CONFIG: dict = {
    "mqtt_host": "",
    "mqtt_port": 1883,
    "mqtt_user": "",
    "mqtt_password": "",
    "autostart": False,
}


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return {**DEFAULT_CONFIG, **json.loads(CONFIG_FILE.read_text())}
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(config: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(config, indent=2))


# ── Autostart ─────────────────────────────────────────────────────────────────

_AUTOSTART_KEY = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
# When frozen by PyInstaller sys.executable is the .exe; otherwise use the script.
_EXE_PATH = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(sys.argv[0])


def enable_autostart() -> None:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{_EXE_PATH}"')
        winreg.CloseKey(key)
    except Exception as exc:
        logging.warning("autostart enable failed: %s", exc)


def disable_autostart() -> None:
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, _AUTOSTART_KEY, 0, winreg.KEY_SET_VALUE)
        winreg.DeleteValue(key, APP_NAME)
        winreg.CloseKey(key)
    except FileNotFoundError:
        pass
    except Exception as exc:
        logging.warning("autostart disable failed: %s", exc)


# ── Camera / microphone detection ─────────────────────────────────────────────

_CAM_STORE = (
    r"SOFTWARE\Microsoft\Windows\CurrentVersion"
    r"\CapabilityAccessManager\ConsentStore"
)


def _check_subkeys_for_active(parent_key) -> bool:
    """Return True if any direct subkey of parent_key has LastUsedTimeStop == 0."""
    i = 0
    while True:
        try:
            name = winreg.EnumKey(parent_key, i)
        except OSError:
            break
        try:
            sub = winreg.OpenKey(parent_key, name)
            try:
                stop, _ = winreg.QueryValueEx(sub, "LastUsedTimeStop")
                if stop == 0:
                    return True
            except FileNotFoundError:
                pass
            finally:
                winreg.CloseKey(sub)
        except OSError:
            pass
        i += 1
    return False


def is_device_in_use(device: str) -> bool:
    """
    Return True if any app currently holds the given device open.

    device: "webcam" | "microphone"

    Windows sets LastUsedTimeStop to 0 while an app is actively using the device.
    Apps are stored under ConsentStore\\{device} as either:
      - Direct subkeys (packaged / Store apps)
      - NonPackaged\\<exe-path> subkeys (Win32 apps)
    """
    base_path = f"{_CAM_STORE}\\{device}"
    try:
        base_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_path)
    except FileNotFoundError:
        return False

    try:
        # Packaged apps sit directly under the device key
        if _check_subkeys_for_active(base_key):
            return True
        # Win32 apps sit under NonPackaged\
        try:
            np_key = winreg.OpenKey(base_key, "NonPackaged")
            result = _check_subkeys_for_active(np_key)
            winreg.CloseKey(np_key)
            return result
        except FileNotFoundError:
            return False
    finally:
        winreg.CloseKey(base_key)


# ── mDNS broker discovery ─────────────────────────────────────────────────────


def discover_mqtt_broker(timeout: float = 5.0) -> tuple[str, int] | None:
    """Scan LAN for an MQTT broker advertised via mDNS. Returns (host, port) or None."""
    found: list[tuple[str, int]] = []

    class _Listener(ServiceListener):
        def add_service(self, zc, type_, name):
            info = zc.get_service_info(type_, name)
            if info:
                addrs = info.parsed_addresses()
                if addrs:
                    found.append((addrs[0], info.port))

        def remove_service(self, zc, type_, name):
            pass

        def update_service(self, zc, type_, name):
            pass

    zc = Zeroconf()
    ServiceBrowser(zc, "_mqtt._tcp.local.", _Listener())
    time.sleep(timeout)
    zc.close()
    return found[0] if found else None


# ── Auto-update ───────────────────────────────────────────────────────────────


def _version_newer(latest: str, current: str) -> bool:
    def parse(v: str):
        return tuple(int(x) for x in v.strip().split("."))
    try:
        return parse(latest) > parse(current)
    except Exception:
        return False


def check_for_update() -> tuple[str, str] | None:
    """Query GitHub releases API. Returns (version, download_url) if newer, else None."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": f"{APP_NAME}/{VERSION}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read())
        latest = data["tag_name"].lstrip("v")
        if _version_newer(latest, VERSION):
            for asset in data.get("assets", []):
                if asset["name"].endswith(".exe"):
                    return latest, asset["browser_download_url"]
    except Exception as exc:
        logging.warning("update check failed: %s", exc)
    return None


def cleanup_old_exe() -> None:
    """Remove ha-agent-old.exe left behind by a previous self-update."""
    if not getattr(sys, "frozen", False):
        return
    old = Path(sys.executable).parent / "ha-agent-old.exe"
    try:
        if old.exists():
            old.unlink()
            logging.info("cleaned up leftover old exe")
    except Exception as exc:
        logging.warning("could not remove old exe: %s", exc)


# ── Setup wizard ──────────────────────────────────────────────────────────────


class SetupDialog:
    """tkinter setup wizard shown on first run or via tray Settings."""

    def __init__(self, config: dict, on_save):
        self._config = dict(config)
        self._on_save = on_save
        self._root = None

    def run(self) -> None:
        self._root = tk.Tk()
        self._root.title(f"{APP_NAME} – Setup")
        self._root.resizable(False, False)
        self._build()
        self._root.mainloop()

    def _build(self) -> None:
        root = self._root
        p = {"padx": 10, "pady": 5}
        f = ttk.Frame(root, padding=16)
        f.grid(sticky="nsew")

        # MQTT host row
        ttk.Label(f, text="MQTT Host:").grid(row=0, column=0, sticky="w", **p)
        self._host = tk.StringVar(value=self._config.get("mqtt_host", ""))
        ttk.Entry(f, textvariable=self._host, width=28).grid(row=0, column=1, **p)
        ttk.Button(f, text="Auto-discover", command=self._discover).grid(row=0, column=2, **p)

        # Port
        ttk.Label(f, text="MQTT Port:").grid(row=1, column=0, sticky="w", **p)
        self._port = tk.StringVar(value=str(self._config.get("mqtt_port", 1883)))
        ttk.Entry(f, textvariable=self._port, width=8).grid(row=1, column=1, sticky="w", **p)

        # Username
        ttk.Label(f, text="Username:").grid(row=2, column=0, sticky="w", **p)
        self._user = tk.StringVar(value=self._config.get("mqtt_user", ""))
        ttk.Entry(f, textvariable=self._user, width=28).grid(row=2, column=1, **p)

        # Password
        ttk.Label(f, text="Password:").grid(row=3, column=0, sticky="w", **p)
        self._password = tk.StringVar(value=self._config.get("mqtt_password", ""))
        ttk.Entry(f, textvariable=self._password, show="*", width=28).grid(row=3, column=1, **p)

        # Autostart
        self._autostart = tk.BooleanVar(value=self._config.get("autostart", False))
        ttk.Checkbutton(
            f, text="Start automatically with Windows", variable=self._autostart
        ).grid(row=4, column=0, columnspan=3, sticky="w", **p)

        # Buttons
        btn = ttk.Frame(f)
        btn.grid(row=5, column=0, columnspan=3, pady=(8, 0))
        ttk.Button(btn, text="Save & Connect", command=self._save).pack(side="left", padx=5)
        ttk.Button(btn, text="Cancel", command=root.destroy).pack(side="left", padx=5)

        # Status line for auto-discover feedback
        self._status = tk.StringVar()
        ttk.Label(f, textvariable=self._status, foreground="#555").grid(
            row=6, column=0, columnspan=3, pady=(4, 0)
        )

    def _discover(self) -> None:
        self._status.set("Scanning for MQTT broker (5 s)…")
        self._root.update()

        def _run():
            result = discover_mqtt_broker(timeout=5.0)
            if result:
                host, port = result
                self._host.set(host)
                self._port.set(str(port))
                self._root.after(0, lambda: self._status.set(f"Found: {host}:{port}"))
            else:
                self._root.after(
                    0, lambda: self._status.set("No MQTT broker found on local network.")
                )

        threading.Thread(target=_run, daemon=True).start()

    def _save(self) -> None:
        host = self._host.get().strip()
        if not host:
            messagebox.showerror("Error", "MQTT host is required.", parent=self._root)
            return
        try:
            port = int(self._port.get())
        except ValueError:
            messagebox.showerror("Error", "Port must be a number.", parent=self._root)
            return

        cfg = {
            "mqtt_host": host,
            "mqtt_port": port,
            "mqtt_user": self._user.get().strip(),
            "mqtt_password": self._password.get(),
            "autostart": self._autostart.get(),
        }
        save_config(cfg)
        if cfg["autostart"]:
            enable_autostart()
        else:
            disable_autostart()

        self._root.destroy()
        self._on_save(cfg)


# ── MQTT Agent ────────────────────────────────────────────────────────────────

_DEVICE_INFO = {
    "identifiers": [f"ha_agent_{HOSTNAME}"],
    "name": HOSTNAME,
    "model": "Windows PC",
    "manufacturer": "HA Agent",
}

_AVAIL_FRAGMENT = {
    "availability_topic": AVAIL_TOPIC,
    "payload_available": "online",
    "payload_not_available": "offline",
    "device": _DEVICE_INFO,
}

_DISCOVERY_CONFIGS: list[tuple[str, dict]] = [
    (
        f"homeassistant/binary_sensor/ha_agent_{HOSTNAME}/camera/config",
        {
            **_AVAIL_FRAGMENT,
            "name": "Camera Active",
            "unique_id": f"{HOSTNAME}_camera",
            "state_topic": CAMERA_TOPIC,
            "payload_on": "ON",
            "payload_off": "OFF",
            "device_class": "running",
        },
    ),
    (
        f"homeassistant/binary_sensor/ha_agent_{HOSTNAME}/microphone/config",
        {
            **_AVAIL_FRAGMENT,
            "name": "Microphone Active",
            "unique_id": f"{HOSTNAME}_microphone",
            "state_topic": MIC_TOPIC,
            "payload_on": "ON",
            "payload_off": "OFF",
            "device_class": "running",
        },
    ),
    (
        f"homeassistant/device_tracker/ha_agent_{HOSTNAME}/config",
        {
            **_AVAIL_FRAGMENT,
            "name": f"{HOSTNAME} PC",
            "unique_id": f"{HOSTNAME}_pc",
            "state_topic": STATUS_TOPIC,
            "payload_home": "home",
            "payload_not_home": "not_home",
            "source_type": "router",
        },
    ),
    (
        f"homeassistant/notify/ha_agent_{HOSTNAME}/config",
        {
            **_AVAIL_FRAGMENT,
            "name": f"{HOSTNAME} Notify",
            "unique_id": f"{HOSTNAME}_notify",
            "command_topic": NOTIFY_TOPIC,
        },
    ),
]


class MQTTAgent:
    def __init__(self, config: dict, notify_callback=None) -> None:
        self._config = config
        self._notify_callback = notify_callback  # callable(title, message)
        self._camera: bool | None = None
        self._mic: bool | None = None
        self._connected = False
        self._stop = threading.Event()
        self._client: mqtt.Client | None = None

    # ── public ────────────────────────────────────────────────────────────────

    def start(self) -> None:
        self._stop.clear()
        client = mqtt.Client(client_id=f"ha-agent-{HOSTNAME}", clean_session=True)
        if self._config.get("mqtt_user"):
            client.username_pw_set(
                self._config["mqtt_user"], self._config.get("mqtt_password", "")
            )
        client.will_set(AVAIL_TOPIC, "offline", retain=True)
        client.on_connect = self._on_connect
        client.on_disconnect = self._on_disconnect
        client.on_message = self._on_message
        client.reconnect_delay_set(min_delay=5, max_delay=60)
        self._client = client

        try:
            client.connect(
                self._config["mqtt_host"],
                int(self._config.get("mqtt_port", 1883)),
                keepalive=60,
            )
        except Exception as exc:
            logging.error("MQTT connect failed: %s", exc)
            return

        client.loop_start()
        threading.Thread(target=self._poll_loop, daemon=True).start()

    def stop(self) -> None:
        self._stop.set()
        if self._client:
            try:
                self._client.publish(AVAIL_TOPIC, "offline", retain=True)
                time.sleep(0.3)
            except Exception:
                pass
            self._client.loop_stop()
            self._client.disconnect()

    @property
    def status_text(self) -> str:
        host = self._config.get("mqtt_host", "?")
        conn = "connected" if self._connected else "disconnected"
        cam = "ON" if self._camera else "OFF"
        mic = "ON" if self._mic else "OFF"
        return f"Broker: {host} ({conn})\nCamera: {cam}    Microphone: {mic}"

    # ── private ───────────────────────────────────────────────────────────────

    def _on_connect(self, client, userdata, flags, rc) -> None:
        if rc == 0:
            self._connected = True
            logging.info("MQTT connected to %s", self._config.get("mqtt_host"))
            self._publish_discovery()
            client.publish(AVAIL_TOPIC, "online", retain=True)
            client.publish(STATUS_TOPIC, "home")
            client.subscribe(NOTIFY_TOPIC)
        else:
            logging.error("MQTT connection refused (rc=%d)", rc)

    def _on_disconnect(self, client, userdata, rc) -> None:
        self._connected = False
        if rc != 0:
            logging.warning("MQTT disconnected unexpectedly (rc=%d)", rc)

    def _on_message(self, client, userdata, msg) -> None:
        if msg.topic != NOTIFY_TOPIC or not self._notify_callback:
            return
        try:
            payload = json.loads(msg.payload.decode())
        except Exception:
            logging.warning("notify: could not parse payload: %r", msg.payload)
            return
        title = str(payload.get("title", APP_NAME))
        message = str(payload.get("message", ""))
        if message:
            logging.info("notify: %s — %s", title, message)
            self._notify_callback(title, message)

    def _publish_discovery(self) -> None:
        for topic, payload in _DISCOVERY_CONFIGS:
            self._client.publish(topic, json.dumps(payload), retain=True)
            logging.debug("discovery published: %s", topic)

    def _poll_loop(self) -> None:
        while not self._stop.is_set():
            if self._connected:
                self._publish_states()
            self._stop.wait(POLL_INTERVAL)

    def _publish_states(self) -> None:
        camera = is_device_in_use("webcam")
        mic = is_device_in_use("microphone")

        if camera != self._camera:
            self._client.publish(CAMERA_TOPIC, "ON" if camera else "OFF")
            self._camera = camera
            logging.debug("camera -> %s", self._camera)

        if mic != self._mic:
            self._client.publish(MIC_TOPIC, "ON" if mic else "OFF")
            self._mic = mic
            logging.debug("microphone -> %s", self._mic)


# ── System tray ───────────────────────────────────────────────────────────────


def _make_tray_icon(connected: bool = True) -> Image.Image:
    color = (34, 139, 34) if connected else (180, 40, 40)
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    ImageDraw.Draw(img).ellipse((4, 4, 60, 60), fill=color)
    return img


class TrayApp:
    def __init__(self, config: dict) -> None:
        self._config = config
        self._icon: pystray.Icon | None = None
        self._agent = MQTTAgent(config, notify_callback=self._show_notification)

    def run(self) -> None:
        self._agent.start()
        self._icon = pystray.Icon(
            APP_NAME,
            _make_tray_icon(),
            APP_NAME,
            self._build_menu(),
        )
        threading.Thread(target=self._check_update_bg, daemon=True).start()
        self._icon.run()

    # ── tray menu actions ─────────────────────────────────────────────────────

    def _build_menu(self) -> pystray.Menu:
        return pystray.Menu(
            pystray.MenuItem("Status", self._show_status),
            pystray.MenuItem("Settings…", self._open_settings),
            pystray.MenuItem("Check for updates", self._check_update_manual),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", self._exit),
        )

    def _show_status(self, icon, item) -> None:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, self._agent.status_text, APP_NAME, 0x40)

    def _show_notification(self, title: str, message: str) -> None:
        if self._icon:
            self._icon.notify(message, title)

    def _open_settings(self, icon, item) -> None:
        def on_save(new_config: dict) -> None:
            self._config = new_config
            self._agent.stop()
            self._agent = MQTTAgent(new_config, notify_callback=self._show_notification)
            self._agent.start()

        threading.Thread(
            target=lambda: SetupDialog(self._config, on_save).run(), daemon=True
        ).start()

    def _check_update_bg(self) -> None:
        time.sleep(5)  # let the icon finish initialising before we show any notification
        self._run_update_check(manual=False)

    def _check_update_manual(self, icon, item) -> None:
        threading.Thread(
            target=lambda: self._run_update_check(manual=True), daemon=True
        ).start()

    def _run_update_check(self, manual: bool = False) -> None:
        result = check_for_update()
        if result:
            new_version, download_url = result
            self._show_notification(APP_NAME, f"Update v{new_version} found, downloading…")
            self._do_update(new_version, download_url)
        elif manual:
            self._show_notification(APP_NAME, f"Already on the latest version ({VERSION}).")

    def _do_update(self, new_version: str, download_url: str) -> None:
        if not getattr(sys, "frozen", False):
            logging.info("running as script — skipping self-replace for v%s", new_version)
            return
        try:
            exe_path = Path(sys.executable)
            tmp_path = exe_path.parent / "ha-agent-new.exe"

            logging.info("downloading v%s from %s", new_version, download_url)
            urllib.request.urlretrieve(download_url, str(tmp_path))

            self._show_notification(APP_NAME, f"v{new_version} ready — restarting…")
            time.sleep(2)

            self._agent.stop()

            # Windows lets you rename a running .exe; swap new into place then restart.
            old_path = exe_path.parent / "ha-agent-old.exe"
            if old_path.exists():
                old_path.unlink()
            exe_path.rename(old_path)
            tmp_path.rename(exe_path)

            subprocess.Popen([str(exe_path)])
            if self._icon:
                self._icon.stop()
        except Exception as exc:
            logging.error("update failed: %s", exc)
            self._show_notification(APP_NAME, f"Update failed: {exc}")

    def _exit(self, icon, item) -> None:
        self._agent.stop()
        icon.stop()


# ── Entry point ───────────────────────────────────────────────────────────────


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
    )

    cleanup_old_exe()

    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--setup", action="store_true", help="Open settings dialog")
    parser.add_argument("--uninstall", action="store_true", help="Remove autostart entry")
    args = parser.parse_args()

    if args.uninstall:
        disable_autostart()
        print("Autostart entry removed.")
        return

    config = load_config()

    if args.setup or not config.get("mqtt_host"):
        saved: dict = {}

        def on_save(cfg: dict) -> None:
            saved.update(cfg)

        SetupDialog(config, on_save).run()

        if not saved.get("mqtt_host"):
            # User cancelled setup
            sys.exit(0)

        config = saved

    TrayApp(config).run()


if __name__ == "__main__":
    main()
