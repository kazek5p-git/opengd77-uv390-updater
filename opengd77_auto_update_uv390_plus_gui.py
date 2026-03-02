#!/usr/bin/env python3
"""GUI launcher for OpenGD77 auto updater (UV390 Plus 10W)."""

from __future__ import annotations

import os
import queue
import re
import threading
import traceback
from pathlib import Path
import sys
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent
os.environ.setdefault("OPENGD77_BASE_DIR", str(APP_DIR))

CORE_IMPORT_ERROR: Exception | None = None
core = None
try:
    import opengd77_auto_update_uv390_plus as core  # type: ignore[assignment]
except Exception as ex:  # pragma: no cover - startup safety path
    CORE_IMPORT_ERROR = ex


class UpdaterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("OpenGD77 Updater - TYT UV390 Plus 10W")
        self.root.geometry("840x560")
        self.root.minsize(760, 500)
        self.root.report_callback_exception = self._report_tk_exception

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.is_busy = False

        self.check_only_var = tk.BooleanVar(value=False)
        self.force_var = tk.BooleanVar(value=False)
        self.dfu_timeout_var = tk.StringVar(value="90")
        self.status_var = tk.StringVar(value="Gotowe do uruchomienia")
        self.last_message_var = tk.StringVar(value="Brak komunikatow.")

        self._build_ui()
        self._bind_shortcuts()
        self.root.protocol("WM_DELETE_WINDOW", self._request_close)
        self.root.after(100, self._flush_logs)
        self.root.after(250, self._announce_startup)

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=12)
        container.pack(fill="both", expand=True)

        heading = ttk.Label(
            container,
            text="OpenGD77 Updater dla TYT UV390 Plus 10W",
        )
        heading.pack(fill="x", pady=(0, 8))

        top = ttk.LabelFrame(container, text="Opcje", padding=10)
        top.pack(fill="x")

        self.check_only_cb = ttk.Checkbutton(
            top,
            text="Tylko sprawdz wersje (bez flashowania)",
            variable=self.check_only_var,
        )
        self.check_only_cb.grid(
            row=0, column=0, sticky="w", padx=(0, 16)
        )
        self.force_cb = ttk.Checkbutton(
            top,
            text="Wymus ponowne flashowanie",
            variable=self.force_var,
        )
        self.force_cb.grid(
            row=0, column=1, sticky="w", padx=(0, 16)
        )

        self.timeout_label = ttk.Label(top, text="Timeout DFU (sekundy):")
        self.timeout_label.grid(row=0, column=2, sticky="e")
        self.timeout_entry = ttk.Entry(top, textvariable=self.dfu_timeout_var, width=6)
        self.timeout_entry.grid(row=0, column=3, sticky="w", padx=(6, 0))

        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)

        actions = ttk.Frame(container)
        actions.pack(fill="x", pady=(10, 8))

        self.start_btn = ttk.Button(actions, text="Start (Alt+S)", command=self.start_update)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(actions, text="Zamknij (Alt+Z)", command=self._request_close)
        self.stop_btn.pack(side="left", padx=(8, 0))

        hints = ttk.Label(
            container,
            text="Skroty: Alt+S Start, Alt+Z Zamknij, Alt+L Log, Alt+T Timeout, Alt+F4 Zamknij, F1 Pomoc.",
        )
        hints.pack(fill="x", pady=(0, 8))

        status_frame = ttk.LabelFrame(container, text="Status", padding=8)
        status_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(status_frame, text="Biezacy status:").grid(row=0, column=0, sticky="w")
        self.status_entry = ttk.Entry(status_frame, textvariable=self.status_var, state="readonly")
        self.status_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ttk.Label(status_frame, text="Ostatni komunikat:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.last_message_entry = ttk.Entry(status_frame, textvariable=self.last_message_var, state="readonly")
        self.last_message_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(6, 0))
        status_frame.columnconfigure(1, weight=1)

        log_frame = ttk.LabelFrame(container, text="Log (Alt+L)")
        log_frame.pack(fill="both", expand=True)

        self.log_box = scrolledtext.ScrolledText(log_frame, wrap="word", height=20)
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)
        self.log_box.bind("<Tab>", self._on_log_tab)
        self.log_box.bind("<Shift-Tab>", self._on_log_shift_tab)
        self.log_box.bind("<ISO_Left_Tab>", self._on_log_shift_tab)
        self.log_box.bind("<Key>", self._block_log_edit)
        self.log_box.bind("<<Paste>>", lambda _e: "break")
        self.log_box.bind("<<Cut>>", lambda _e: "break")

    def _bind_shortcuts(self) -> None:
        self.root.bind_all("<Alt-s>", self._on_alt_start)
        self.root.bind_all("<Alt-z>", self._on_alt_close)
        self.root.bind_all("<Alt-F4>", self._on_alt_close)
        self.root.bind_all("<Alt-l>", self._on_alt_log)
        self.root.bind_all("<Alt-t>", self._on_alt_timeout)
        self.root.bind_all("<F1>", self._on_help)

    def _on_alt_start(self, _event=None):
        if self.start_btn["state"] != "disabled":
            self.start_btn.invoke()
        return "break"

    def _on_alt_close(self, _event=None):
        self._request_close()
        return "break"

    def _on_alt_log(self, _event=None):
        self.log_box.focus_set()
        return "break"

    def _on_alt_timeout(self, _event=None):
        self.timeout_entry.focus_set()
        self.timeout_entry.selection_range(0, "end")
        return "break"

    def _on_help(self, _event=None):
        self._show_info(
            "Pomoc",
            "Sterowanie:\n"
            "- Alt+S: Start aktualizacji\n"
            "- Alt+Z: Zamknij\n"
            "- Alt+F4: Zamknij (z potwierdzeniem)\n"
            "- Alt+L: Przejdz do logu\n"
            "- Alt+T: Przejdz do pola Timeout\n"
            "- F1: Ta pomoc",
        )
        return "break"

    def _on_log_tab(self, _event=None):
        nxt = self.log_box.tk_focusNext()
        if nxt is not None:
            nxt.focus_set()
        return "break"

    def _on_log_shift_tab(self, _event=None):
        prev = self.log_box.tk_focusPrev()
        if prev is not None:
            prev.focus_set()
        return "break"

    def _request_close(self) -> None:
        update_running = self.is_busy or (self.worker is not None and self.worker.is_alive())
        if update_running:
            messagebox.showwarning(
                "Aktualizacja w toku",
                "Nie mozna zamknac programu podczas aktualizacji radia.",
                parent=self.root,
            )
            return

        close_ok = messagebox.askyesno(
            "Potwierdzenie zamkniecia",
            "Czy na pewno chcesz zamknac program?",
            default="no",
            parent=self.root,
        )
        if close_ok:
            self.root.destroy()

    def _announce_startup(self) -> None:
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(250, lambda: self.root.attributes("-topmost", False))
        except Exception:
            pass
        self.start_btn.focus_set()
        self._show_info(
            "Aplikacja gotowa",
            "OpenGD77 Updater zostal uruchomiony.\n"
            "Uzyj Tab aby przechodzic po kontrolkach.\n"
            "Nacisnij Alt+S, aby rozpoczac.",
        )

    def _block_log_edit(self, event: tk.Event) -> str | None:
        ctrl = bool(getattr(event, "state", 0) & 0x4)
        key = (getattr(event, "keysym", "") or "").lower()
        if ctrl and key in {"c", "a"}:
            return None
        if key in {
            "up",
            "down",
            "left",
            "right",
            "prior",
            "next",
            "home",
            "end",
            "escape",
        }:
            return None
        if key in {"tab", "iso_left_tab"}:
            return "break"
        return "break"

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)
        self.last_message_var.set(text)

    def _report_tk_exception(self, exc_type, exc_value, exc_tb) -> None:
        details = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        msg = f"Wystapil nieoczekiwany blad:\n{exc_value}"
        try:
            log_file = APP_DIR / "OpenGD77_GUI_error.log"
            with open(log_file, "w", encoding="utf-8") as f:
                f.write(details)
            msg += f"\n\nSzczegoly zapisano do:\n{log_file}"
        except Exception:
            pass
        messagebox.showerror("Blad aplikacji", msg, parent=self.root)

    def log(self, msg: str) -> None:
        self.log_queue.put(msg)

    def _flush_logs(self) -> None:
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_box.insert("end", msg + "\n")
                self.log_box.see("end")
                if msg.strip():
                    self.last_message_var.set(msg.strip())
        except queue.Empty:
            pass
        self.root.after(100, self._flush_logs)

    def _set_busy(self, busy: bool, status: str) -> None:
        self.is_busy = busy
        self._set_status(status)
        if busy:
            self.start_btn.configure(state="disabled")
            self.timeout_entry.configure(state="disabled")
            self.stop_btn.configure(state="disabled")
        else:
            self.start_btn.configure(state="normal")
            self.timeout_entry.configure(state="normal")
            self.stop_btn.configure(state="normal")

    def _ask_yes_no(self, title: str, prompt: str, default_no: bool = True) -> bool:
        event = threading.Event()
        result = {"answer": False}

        def ask() -> None:
            options = {"title": title, "message": prompt, "parent": self.root}
            options["default"] = "no" if default_no else "yes"
            result["answer"] = messagebox.askyesno(**options)
            event.set()

        self.root.after(0, ask)
        event.wait()
        return bool(result["answer"])

    def _ask_ok_cancel(self, title: str, prompt: str) -> bool:
        event = threading.Event()
        result = {"answer": False}

        def ask() -> None:
            result["answer"] = messagebox.askokcancel(title=title, message=prompt, parent=self.root)
            event.set()

        self.root.after(0, ask)
        event.wait()
        return bool(result["answer"])

    def _show_info(self, title: str, prompt: str) -> None:
        self.root.after(0, lambda: messagebox.showinfo(title=title, message=prompt, parent=self.root))

    def _show_error(self, title: str, prompt: str) -> None:
        self.root.after(0, lambda: messagebox.showerror(title=title, message=prompt, parent=self.root))

    def start_update(self) -> None:
        if self.worker and self.worker.is_alive():
            return

        if core is None:
            self._show_error(
                "Blad startu",
                "Nie mozna zaladowac modulu aktualizatora.\n"
                f"Szczegoly: {CORE_IMPORT_ERROR}",
            )
            return

        try:
            timeout_s = int(self.dfu_timeout_var.get().strip())
            if timeout_s <= 0:
                raise ValueError
        except Exception:
            self._show_error("Blad", "Timeout DFU musi byc dodatnia liczba calkowita.")
            return

        check_only = self.check_only_var.get()
        force = self.force_var.get()

        self._set_busy(True, "Trwa sprawdzanie...")
        self.log("")
        self.log("=== OpenGD77 Auto Updater (GUI) ===")
        self.worker = threading.Thread(
            target=self._run_update_worker,
            args=(check_only, force, timeout_s),
            daemon=True,
        )
        self.worker.start()

    def _run_update_worker(self, check_only: bool, force: bool, timeout_s: int) -> None:
        exit_code = 0
        try:
            core.ensure_dir(core.DOWNLOADS_DIR)

            self.log("Sprawdzam najnowszy release...")
            latest_release = core.get_latest_release_tag()
            latest_date = latest_release[1:]
            self.log(f"Najnowszy release: {latest_release}")

            current_info = None
            com_port = core.find_serial_port()
            if com_port:
                try:
                    current_info = core.read_radio_info(com_port)
                except Exception:
                    current_info = None

            if current_info:
                self.log(
                    f"Radio teraz: git={current_info['git_revision']} build={current_info['build_datetime']} port={com_port}"
                )
            else:
                self.log("Nie udalo sie odczytac biezacej wersji z radia (to nie blokuje aktualizacji).")

            up_to_date = False
            if current_info and re.fullmatch(r"\d{14}", current_info["build_datetime"]):
                build_date = current_info["build_datetime"][:8]
                up_to_date = build_date >= latest_date

            do_flash = not check_only
            if up_to_date and not force:
                self.log("Radio wyglada na aktualne wedlug daty buildu.")
                if check_only:
                    do_flash = False
                else:
                    do_flash = self._ask_yes_no("OpenGD77", "Radio jest aktualne.\nWymusic ponowne flashowanie?")
                    if not do_flash:
                        self.log("Przerwano przez uzytkownika.")

            if not do_flash:
                if check_only:
                    self.log("Tryb tylko sprawdzania zakonczony.")
                exit_code = 0
                return

            self.log("Przygotowuje pliki firmware i narzedzia...")
            firmware_bin = core.ensure_firmware_bin(latest_release)
            donor_bin = core.ensure_donor_bin()
            loader_path = core.ensure_loader_script(latest_release)
            libusb_dll = core.ensure_libusb_dll()

            self.log("")
            self.log("Wprowadz radio w tryb aktualizacji (DFU):")
            self.log("1) Wylacz radio.")
            self.log("2) Przytrzymaj PTT + S1 (gorny boczny przycisk).")
            self.log("3) Wlacz radio (ekran powinien byc czarny).")
            self.log("4) Podlacz USB.")

            ready = self._ask_ok_cancel(
                "DFU",
                "Ustaw radio w DFU:\n"
                "1) Wylacz radio\n"
                "2) Przytrzymaj PTT + S1\n"
                "3) Wlacz radio (czarny ekran)\n"
                "4) Podlacz USB\n\n"
                "Kliknij OK, gdy jestes gotowy.",
            )
            if not ready:
                self.log("Przerwano przed oczekiwaniem na DFU.")
                exit_code = 0
                return

            self.log("Czekam na DFU...")
            if not core.wait_for_dfu(libusb_dll, timeout_s=timeout_s):
                self.log("Nie wykryto DFU. Sprawdz tryb radia oraz sterownik WinUSB dla VID_0483&PID_DF11.")
                driver_inf = core.find_winusb_driver_inf()
                if driver_inf and self._ask_yes_no(
                    "Brak DFU", "Nie wykryto DFU.\nSprobowac automatycznie zainstalowac sterownik WinUSB?"
                ):
                    self.log(f"Instalacja sterownika z: {driver_inf}")
                    ok, out = core.install_winusb_driver(driver_inf, elevate=True)
                    if out:
                        self.log(out)
                    if ok:
                        ready_retry = self._ask_ok_cancel(
                            "DFU",
                            "Sterownik zainstalowany.\n"
                            "Ustaw ponownie radio w DFU (PTT + S1) i kliknij OK, aby kontynuowac.",
                        )
                        if not ready_retry:
                            self.log("Przerwano po instalacji sterownika.")
                            exit_code = 0
                            return
                        self.log("Ponowne oczekiwanie na DFU...")
                        if not core.wait_for_dfu(libusb_dll, timeout_s=timeout_s):
                            self._show_error("Brak DFU", "Dalej nie wykryto DFU po instalacji sterownika.")
                            exit_code = 4
                            return
                    else:
                        self._show_error(
                            "Blad sterownika",
                            "Nie udalo sie automatycznie zainstalowac sterownika WinUSB.",
                        )
                        exit_code = 4
                        return
                else:
                    self._show_error("Brak DFU", "Nie wykryto DFU. Sprawdz radio i sterownik WinUSB.")
                    exit_code = 4
                    return

            self.log("DFU wykryte.")
            if not self._ask_yes_no("OpenGD77", "Rozpoczac flashowanie teraz?"):
                self.log("Przerwano przed flashowaniem.")
                exit_code = 0
                return

            flash_args = ["-s", str(donor_bin), "-m", "MD-UV380", "-f", str(firmware_bin)]
            self.log("Start flashowania...")
            rc = core.run_loader_with_backend(loader_path, libusb_dll, flash_args)
            if rc != 0:
                self.log(f"Flash zakonczyl sie bledem, kod={rc}")
                self._show_error("Blad flashowania", f"Flash zakonczyl sie bledem (kod {rc}).")
                exit_code = rc
                return

            self.log("Flash zakonczony.")
            self.log("Poczekaj 5-10 sekund, radio powinno wrocic do normalnego trybu.")

            com_after = core.find_serial_port()
            if com_after:
                try:
                    info_after = core.read_radio_info(com_after)
                    if info_after:
                        self.log(
                            f"Po update: git={info_after['git_revision']} build={info_after['build_datetime']} port={com_after}"
                        )
                except Exception:
                    pass

            self.log("Gotowe.")
            self._show_info("Sukces", "Aktualizacja zakonczona.")

        except Exception as ex:
            self.log(f"Blad: {ex}")
            self.log(traceback.format_exc())
            self._show_error("Blad", f"Wystapil blad:\n{ex}")
            exit_code = 10
        finally:
            final_status = "Zakonczono" if exit_code == 0 else f"Blad (kod {exit_code})"
            self.root.after(0, lambda: self._set_busy(False, final_status))


def main() -> int:
    if core is None:
        messagebox.showerror(
            "Blad startu",
            "Nie mozna uruchomic OpenGD77 Updater.\n"
            f"Blad importu: {CORE_IMPORT_ERROR}\n\n"
            f"Katalog aplikacji: {APP_DIR}",
        )
        return 2

    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    UpdaterApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as ex:  # pragma: no cover - startup safety path
        details = traceback.format_exc()
        try:
            log_file = APP_DIR / "OpenGD77_GUI_error.log"
            with open(log_file, "w", encoding="utf-8") as f:
                f.write(details)
            extra = f"\n\nSzczegoly zapisano do:\n{log_file}"
        except Exception:
            extra = ""
        messagebox.showerror("Blad krytyczny", f"Aplikacja zatrzymana: {ex}{extra}")
        raise
