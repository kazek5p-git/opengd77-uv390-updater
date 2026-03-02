#!/usr/bin/env python3
"""
OpenGD77 auto-updater for TYT UV390 Plus 10W (MD-UV380 family).

Flow:
1) Check currently installed firmware (if radio is in normal COM mode).
2) Check latest OpenGD77 release from opengd77.com.
3) Download needed files (firmware zip + donor firmware).
4) Ask user to put radio in update mode (PTT + S1), wait for DFU.
5) Flash firmware automatically using OpenGD77 STM32 loader.
"""

from __future__ import annotations

import argparse
import os
import re
import runpy
import struct
import subprocess
import sys
import time
import urllib.request
import zipfile
from pathlib import Path


def resolve_base_dir() -> Path:
    # Allow wrappers (GUI / packaged exe) to override storage location.
    env_base = os.environ.get("OPENGD77_BASE_DIR")
    if env_base:
        return Path(env_base).expanduser().resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolve_runtime_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(str(sys._MEIPASS)).resolve()
    return Path(__file__).resolve().parent


BASE_DIR = resolve_base_dir()
RUNTIME_DIR = resolve_runtime_dir()
DOWNLOADS_DIR = BASE_DIR / "downloads"
FIRMWARE_WORK_DIR = BASE_DIR / "firmware_extracted"
DONOR_WORK_DIR = BASE_DIR / "donor_extracted"
TOOLS_DIR = BASE_DIR / "stm32_loader_tools"
LIBUSB_DIR = BASE_DIR / "libusb"
LIBUSB_DLL = LIBUSB_DIR / "dll64" / "MinGW64" / "dll" / "libusb-1.0.dll"

RELEASES_ROOT = "https://www.opengd77.com/downloads/releases/MDUV380_DM1701/"
FIRMWARE_FILE_NAME = "OpenMDUV380_10W_PLUS.zip"
DONOR_URL = "https://www.passion-radio.com/index.php?controller=attachment&id_attachment=760"
DONOR_BIN_NAME = "MD9600-CSV(2571V5)-V26.45.bin"
LIBUSB_7Z_URL = "https://github.com/libusb/libusb/releases/download/v1.0.29/libusb-1.0.29.7z"

DFU_VID = 0x0483
DFU_PID = 0xDF11

SERIAL_VID = 0x1FC9
SERIAL_PID = 0x0094


BUNDLED_RESOURCES_DIR = RUNTIME_DIR / "resources"
BUNDLED_LOADER = BUNDLED_RESOURCES_DIR / "opengd77_stm32_firmware_loader.py"
BUNDLED_LIBUSB_DLL = BUNDLED_RESOURCES_DIR / "libusb-1.0.dll"
BUNDLED_DONOR_BIN = BUNDLED_RESOURCES_DIR / DONOR_BIN_NAME
BUNDLED_DRIVER_INF = BUNDLED_RESOURCES_DIR / "driver" / "usb_device.inf"
BUNDLED_DRIVER_CAT = BUNDLED_RESOURCES_DIR / "driver" / "usb_device.cat"


def log(msg: str) -> None:
    print(msg, flush=True)


def yes_no(prompt: str, default_no: bool = True) -> bool:
    suffix = "[t/N]" if default_no else "[T/n]"
    ans = input(f"{prompt} {suffix}: ").strip().lower()
    if not ans:
        return not default_no
    return ans in {"t", "tak", "y", "yes"}


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def download_file(url: str, dest: Path) -> None:
    ensure_dir(dest.parent)
    if dest.exists() and dest.stat().st_size > 0:
        return
    log(f"Pobieranie: {url}")
    with urllib.request.urlopen(url) as resp, open(dest, "wb") as out:
        out.write(resp.read())


def get_text(url: str) -> str:
    with urllib.request.urlopen(url) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def get_latest_release_tag() -> str:
    html = get_text(RELEASES_ROOT)
    tags = sorted(set(re.findall(r"R\d{8}", html)))
    if not tags:
        raise RuntimeError("Nie udalo sie pobrac listy release z OpenGD77.")
    return tags[-1]


def find_serial_port() -> str | None:
    try:
        from serial.tools import list_ports
    except Exception:
        return None

    for p in list_ports.comports():
        if (p.vid == SERIAL_VID and p.pid == SERIAL_PID) or "VID_1FC9&PID_0094" in (p.hwid or ""):
            return p.device
    return None


def read_radio_info(port: str) -> dict | None:
    import serial

    ser = serial.Serial(
        port=port,
        baudrate=115200,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        timeout=1.0,
        write_timeout=1.0,
    )
    try:
        # Enter data mode.
        ser.reset_input_buffer()
        ser.write(bytes([ord("C"), 254]))
        time.sleep(0.15)
        _ = ser.read(ser.in_waiting or 1)

        # Read radio info packet.
        req = bytes([ord("R"), 9, 0, 0, 0, 0, 0, 8])
        ser.reset_input_buffer()
        ser.write(req)

        raw = b""
        for _ in range(25):
            n = ser.in_waiting
            if n:
                raw += ser.read(n)
                if len(raw) >= 3:
                    total = (raw[1] << 8) | raw[2]
                    if len(raw) >= 3 + total:
                        break
            time.sleep(0.05)

        if len(raw) < 3 or raw[0] != ord("R"):
            return None

        total = (raw[1] << 8) | raw[2]
        payload = raw[3 : 3 + total]
        if len(payload) < 46:
            return None

        struct_version, radio_type, git_rev, build_dt, flash_id, features = struct.unpack(
            "<II16s16sIH", payload[:46]
        )
        return {
            "struct_version": struct_version,
            "radio_type": radio_type,
            "git_revision": git_rev.split(b"\x00", 1)[0].decode("ascii", errors="ignore"),
            "build_datetime": build_dt.split(b"\x00", 1)[0].decode("ascii", errors="ignore"),
            "flash_id": flash_id,
            "features": features,
        }
    finally:
        try:
            ser.write(bytes([ord("C"), 7]))
        except Exception:
            pass
        ser.close()


def ensure_firmware_bin(release_tag: str) -> Path:
    zip_url = f"{RELEASES_ROOT}{release_tag}/firmware/{FIRMWARE_FILE_NAME}"
    zip_path = DOWNLOADS_DIR / f"{FIRMWARE_FILE_NAME[:-4]}_{release_tag}.zip"
    download_file(zip_url, zip_path)

    ensure_dir(FIRMWARE_WORK_DIR)
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(FIRMWARE_WORK_DIR)

    bin_path = FIRMWARE_WORK_DIR / "OpenMDUV380_10W_PLUS.bin"
    if not bin_path.exists():
        raise RuntimeError(f"Brak pliku firmware bin po rozpakowaniu: {bin_path}")
    return bin_path


def ensure_donor_bin() -> Path:
    if BUNDLED_DONOR_BIN.exists():
        return BUNDLED_DONOR_BIN

    donor_zip = DOWNLOADS_DIR / "Donor_MD9600_V26.45.zip"
    download_file(DONOR_URL, donor_zip)
    ensure_dir(DONOR_WORK_DIR)
    with zipfile.ZipFile(donor_zip, "r") as zf:
        zf.extractall(DONOR_WORK_DIR)
    donor_bin = DONOR_WORK_DIR / DONOR_BIN_NAME
    if not donor_bin.exists():
        raise RuntimeError(f"Brak donor bin: {donor_bin}")
    return donor_bin


def find_loader_script() -> Path | None:
    if BUNDLED_LOADER.exists():
        return BUNDLED_LOADER

    matches = list(BASE_DIR.rglob("opengd77_stm32_firmware_loader.py"))
    return matches[0] if matches else None


def ensure_loader_script(release_tag: str) -> Path:
    existing = find_loader_script()
    if existing:
        return existing

    sources_url = f"{RELEASES_ROOT}{release_tag}/sources_and_tools/"
    html = get_text(sources_url)
    m = re.search(r"(OpenGD77_MDUV380_DM1701_\d+\.zip)", html)
    if not m:
        raise RuntimeError("Nie znaleziono paczki tools w release.")

    tools_zip_name = m.group(1)
    tools_zip_url = f"{sources_url}{tools_zip_name}"
    tools_zip_path = DOWNLOADS_DIR / tools_zip_name
    download_file(tools_zip_url, tools_zip_path)

    ensure_dir(TOOLS_DIR)
    with zipfile.ZipFile(tools_zip_path, "r") as zf:
        wanted = [
            n
            for n in zf.namelist()
            if n.endswith("MDUV380_firmware/tools/opengd77_stm32_firmware_loader.py")
        ]
        if not wanted:
            raise RuntimeError("Brak opengd77_stm32_firmware_loader.py w paczce tools.")
        zf.extract(wanted[0], TOOLS_DIR)

    loader = find_loader_script()
    if not loader:
        raise RuntimeError("Nie udalo sie przygotowac loadera STM32.")
    return loader


def ensure_libusb_dll() -> Path:
    if BUNDLED_LIBUSB_DLL.exists():
        return BUNDLED_LIBUSB_DLL

    if LIBUSB_DLL.exists():
        return LIBUSB_DLL

    ensure_dir(LIBUSB_DIR)
    archive = LIBUSB_DIR / "libusb-1.0.29.7z"
    download_file(LIBUSB_7Z_URL, archive)

    seven_zip = shutil_which("7z")
    if not seven_zip:
        raise RuntimeError("Brak 7z.exe. Zainstaluj 7-Zip i uruchom ponownie.")

    ensure_dir(LIBUSB_DLL.parent)
    cmd = [
        seven_zip,
        "x",
        str(archive),
        f"-o{LIBUSB_DIR / 'dll64'}",
        "-y",
        "MinGW64\\dll\\libusb-1.0.dll",
    ]
    subprocess.check_call(cmd)

    if not LIBUSB_DLL.exists():
        raise RuntimeError("Nie udalo sie wypakowac libusb-1.0.dll.")
    return LIBUSB_DLL


def shutil_which(exe: str) -> str | None:
    for p in os.environ.get("PATH", "").split(os.pathsep):
        cand = Path(p) / exe
        if cand.exists():
            return str(cand)
    return None


def find_winusb_driver_inf() -> Path | None:
    candidates = [
        BUNDLED_DRIVER_INF,
        BASE_DIR / "wdi_driver_extract" / "usb_device.inf",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def install_winusb_driver(inf_path: Path, elevate: bool = True) -> tuple[bool, str]:
    if sys.platform != "win32":
        return False, "Instalacja sterownika automatycznie jest dostepna tylko na Windows."
    if not inf_path.exists():
        return False, f"Brak pliku INF: {inf_path}"

    if elevate:
        # Run PnPUtil with UAC prompt and wait for completion.
        arg = f'/add-driver "{inf_path}" /install'
        ps = (
            "$p = Start-Process -FilePath 'pnputil.exe' "
            f"-ArgumentList '{arg}' -Verb RunAs -PassThru -Wait; exit $p.ExitCode"
        )
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True,
            text=True,
        )
    else:
        proc = subprocess.run(
            ["pnputil.exe", "/add-driver", str(inf_path), "/install"],
            capture_output=True,
            text=True,
        )

    output = (proc.stdout or "") + (proc.stderr or "")
    return (proc.returncode == 0), output.strip()


def list_dfu_devices(libusb_dll: Path):
    import usb.backend.libusb1
    import usb.core

    backend = usb.backend.libusb1.get_backend(find_library=lambda _name: str(libusb_dll))
    if backend is None:
        raise RuntimeError("Nie udalo sie uruchomic backendu libusb.")
    return list(usb.core.find(find_all=True, idVendor=DFU_VID, idProduct=DFU_PID, backend=backend))


def wait_for_dfu(libusb_dll: Path, timeout_s: int = 90) -> bool:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            if list_dfu_devices(libusb_dll):
                return True
        except Exception:
            pass
        time.sleep(1)
    return False


def run_loader_with_backend(loader_path: Path, libusb_dll: Path, args: list[str]) -> int:
    import usb.backend.libusb1
    import usb.core

    backend = usb.backend.libusb1.get_backend(find_library=lambda _name: str(libusb_dll))
    if backend is None:
        raise RuntimeError("Brak backendu libusb.")

    original_find = usb.core.find
    argv_backup = sys.argv[:]

    def patched_find(*p_args, **p_kwargs):
        p_kwargs.setdefault("backend", backend)
        return original_find(*p_args, **p_kwargs)

    usb.core.find = patched_find
    sys.argv = [str(loader_path)] + args
    try:
        runpy.run_path(str(loader_path), run_name="__main__")
        return 0
    except SystemExit as ex:
        code = ex.code
        return int(code) if isinstance(code, int) else 0
    finally:
        usb.core.find = original_find
        sys.argv = argv_backup


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="OpenGD77 auto updater for UV390 Plus 10W")
    p.add_argument("--check-only", action="store_true", help="Tylko sprawdz wersje, bez flashowania")
    p.add_argument("--force", action="store_true", help="Wymus aktualizacje nawet jesli radio jest aktualne")
    p.add_argument("--dfu-timeout", type=int, default=90, help="Czas oczekiwania na DFU (sekundy)")
    p.add_argument(
        "--auto-confirm",
        action="store_true",
        help="Tryb bez pytan interaktywnych (dla zewnetrznego GUI).",
    )
    p.add_argument(
        "--auto-driver-install",
        action="store_true",
        help="W trybie auto probuje od razu instalacji sterownika WinUSB po braku DFU.",
    )
    return p.parse_args()


def main() -> int:
    args = parse_args()

    log("=== OpenGD77 Auto Updater (UV390 Plus 10W) ===")
    ensure_dir(DOWNLOADS_DIR)

    try:
        latest_release = get_latest_release_tag()
    except Exception as ex:
        log(f"Blad pobierania latest release: {ex}")
        return 2

    latest_date = latest_release[1:]
    log(f"Najnowszy release: {latest_release}")

    current_info = None
    com_port = find_serial_port()
    if com_port:
        try:
            current_info = read_radio_info(com_port)
        except Exception:
            current_info = None

    if current_info:
        log(
            f"Radio teraz: git={current_info['git_revision']} build={current_info['build_datetime']} port={com_port}"
        )
    else:
        log("Nie udalo sie odczytac biezacej wersji z radia (to nie blokuje aktualizacji).")

    up_to_date = False
    if current_info and re.fullmatch(r"\d{14}", current_info["build_datetime"]):
        build_date = current_info["build_datetime"][:8]
        up_to_date = build_date >= latest_date

    if up_to_date and not args.force:
        log("Radio wyglada na aktualne wedlug daty buildu.")
        if args.check_only:
            return 0
        if args.auto_confirm:
            log("Tryb auto-confirm: wymuszam ponowne flashowanie.")
        elif not yes_no("Wymusic ponowne flashowanie?"):
            log("Przerwano.")
            return 0

    if args.check_only:
        return 0

    try:
        firmware_bin = ensure_firmware_bin(latest_release)
        donor_bin = ensure_donor_bin()
        loader_path = ensure_loader_script(latest_release)
        libusb_dll = ensure_libusb_dll()
    except Exception as ex:
        log(f"Blad przygotowania plikow: {ex}")
        return 3

    log("")
    log("Wprowadz radio w tryb aktualizacji (DFU):")
    log("1) Wylacz radio.")
    log("2) Przytrzymaj PTT + S1 (gorny boczny przycisk).")
    log("3) Wlacz radio (ekran powinien byc czarny).")
    log("4) Podlacz USB.")
    if args.auto_confirm:
        log("Tryb auto-confirm: pomijam pytanie o Enter.")
    else:
        input("Nacisnij Enter, gdy jestes gotowy...")

    log("Czekam na DFU...")
    if not wait_for_dfu(libusb_dll, timeout_s=args.dfu_timeout):
        log("Nie wykryto DFU. Sprawdz tryb radia oraz sterownik WinUSB dla VID_0483&PID_DF11.")
        driver_inf = find_winusb_driver_inf()
        should_try_driver = False
        if driver_inf:
            if args.auto_driver_install:
                should_try_driver = True
            elif not args.auto_confirm:
                should_try_driver = yes_no("Sprobowac automatycznie zainstalowac sterownik WinUSB DFU?")

        if driver_inf and should_try_driver:
            ok, out = install_winusb_driver(driver_inf, elevate=True)
            if out:
                log(out)
            if ok:
                log("Powtorz wejscie w DFU i nacisnij Enter, aby sprawdzic ponownie.")
                if args.auto_confirm:
                    log("Tryb auto-confirm: pomijam pytanie o Enter po instalacji sterownika.")
                else:
                    input("Nacisnij Enter, gdy radio jest ponownie w DFU...")
                log("Ponowne oczekiwanie na DFU...")
                if not wait_for_dfu(libusb_dll, timeout_s=args.dfu_timeout):
                    log("Dalej brak DFU po instalacji sterownika.")
                    return 4
            else:
                log("Nie udalo sie zainstalowac sterownika automatycznie.")
                return 4
        else:
            return 4

    log("DFU wykryte.")
    if args.auto_confirm:
        log("Tryb auto-confirm: rozpoczynam flashowanie bez dodatkowego pytania.")
    elif not yes_no("Rozpoczac flashowanie teraz?"):
        log("Przerwano przed flashowaniem.")
        return 0

    flash_args = [
        "-s",
        str(donor_bin),
        "-m",
        "MD-UV380",
        "-f",
        str(firmware_bin),
    ]

    log("Start flashowania...")
    rc = run_loader_with_backend(loader_path, libusb_dll, flash_args)
    if rc != 0:
        log(f"Flash zakonczyl sie bledem, kod={rc}")
        return rc

    log("Flash zakonczony.")
    log("Poczekaj 5-10 sekund, radio powinno wrocic do normalnego trybu.")
    time.sleep(6)

    com_after = find_serial_port()
    if com_after:
        try:
            info_after = read_radio_info(com_after)
            if info_after:
                log(
                    f"Po update: git={info_after['git_revision']} build={info_after['build_datetime']} port={com_after}"
                )
        except Exception:
            pass

    log("Gotowe.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
