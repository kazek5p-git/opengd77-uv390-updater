#!/usr/bin/env python3
"""Accessible Qt GUI for OpenGD77 updater (UV390 Plus 10W)."""

from __future__ import annotations

import os
import queue
import re
import sys
import threading
import traceback
from dataclasses import dataclass, field
from pathlib import Path

from PySide6.QtCore import QObject, Qt, QTimer, Signal, Slot
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)


if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent
os.environ.setdefault("OPENGD77_BASE_DIR", str(APP_DIR))
os.environ.setdefault("QT_ACCESSIBILITY", "1")

CORE_IMPORT_ERROR: Exception | None = None
core = None
try:
    import opengd77_auto_update_uv390_plus as core  # type: ignore[assignment]
except Exception as ex:  # pragma: no cover
    CORE_IMPORT_ERROR = ex


@dataclass
class PromptRequest:
    title: str
    text: str
    default_no: bool = True
    event: threading.Event = field(default_factory=threading.Event)
    result: bool = False


class UpdaterWindow(QMainWindow):
    signal_log = Signal(str)
    signal_status = Signal(str)
    signal_show_info = Signal(str, str)
    signal_show_error = Signal(str, str)
    signal_ask_yes_no = Signal(object)
    signal_ask_ok_cancel = Signal(object)

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("OpenGD77 Updater - TYT UV390 Plus 10W")
        self.resize(920, 620)
        self.setMinimumSize(820, 520)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.is_busy = False

        self._build_ui()
        self._connect_signals()
        self._bind_shortcuts()
        self._announce_startup()

        self.flush_timer = QTimer(self)
        self.flush_timer.timeout.connect(self._flush_logs)
        self.flush_timer.start(100)

    def _build_ui(self) -> None:
        root = QWidget(self)
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        title = QLabel("OpenGD77 Updater dla TYT UV390 Plus 10W", self)
        title.setAccessibleName("Tytul aplikacji")
        layout.addWidget(title)

        options_group = QGroupBox("Opcje", self)
        options_layout = QGridLayout(options_group)

        self.check_only = QCheckBox("Tylko sprawdz wersje (bez flashowania)", self)
        self.check_only.setAccessibleName("Tryb tylko sprawdzania")
        self.check_only.setAccessibleDescription("Gdy zaznaczone, program sprawdza wersje bez flashowania.")
        options_layout.addWidget(self.check_only, 0, 0, 1, 2)

        self.force = QCheckBox("Wymus ponowne flashowanie", self)
        self.force.setAccessibleName("Wymus ponowne flashowanie")
        self.force.setAccessibleDescription("Wymusza flashowanie nawet gdy wersja jest aktualna.")
        options_layout.addWidget(self.force, 1, 0, 1, 2)

        timeout_label = QLabel("Timeout DFU (sekundy):", self)
        timeout_label.setAccessibleName("Etykieta timeout DFU")
        self.timeout = QSpinBox(self)
        self.timeout.setRange(5, 900)
        self.timeout.setValue(90)
        self.timeout.setAccessibleName("Timeout DFU")
        self.timeout.setAccessibleDescription("Czas oczekiwania na wykrycie DFU.")
        options_layout.addWidget(timeout_label, 2, 0)
        options_layout.addWidget(self.timeout, 2, 1)
        options_layout.setColumnStretch(0, 1)
        layout.addWidget(options_group)

        actions_layout = QHBoxLayout()
        self.start_btn = QPushButton("Start (Alt+S)", self)
        self.start_btn.setAccessibleName("Start aktualizacji")
        self.start_btn.clicked.connect(self.start_update)
        actions_layout.addWidget(self.start_btn)

        self.close_btn = QPushButton("Zamknij (Alt+Z)", self)
        self.close_btn.setAccessibleName("Zamknij program")
        self.close_btn.clicked.connect(self.request_close)
        actions_layout.addWidget(self.close_btn)
        actions_layout.addStretch(1)
        layout.addLayout(actions_layout)

        self.hints = QLabel(
            "Skroty: Alt+S Start, Alt+Z Zamknij, Alt+L Log, Alt+T Timeout, Alt+F4 Zamknij, F1 Pomoc.",
            self,
        )
        self.hints.setAccessibleName("Podpowiedzi skrotow")
        layout.addWidget(self.hints)

        status_group = QGroupBox("Status", self)
        status_layout = QFormLayout(status_group)
        self.status = QLineEdit("Gotowe do uruchomienia", self)
        self.status.setReadOnly(True)
        self.status.setAccessibleName("Biezacy status")
        status_layout.addRow("Biezacy status:", self.status)

        self.last_msg = QLineEdit("Brak komunikatow.", self)
        self.last_msg.setReadOnly(True)
        self.last_msg.setAccessibleName("Ostatni komunikat")
        status_layout.addRow("Ostatni komunikat:", self.last_msg)
        layout.addWidget(status_group)

        log_group = QGroupBox("Log", self)
        log_layout = QVBoxLayout(log_group)
        self.log_box = QPlainTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setTabChangesFocus(True)
        self.log_box.setAccessibleName("Log aktualizacji")
        self.log_box.setAccessibleDescription("Pole tylko do odczytu z przebiegiem aktualizacji.")
        log_layout.addWidget(self.log_box)
        layout.addWidget(log_group, 1)

    def _connect_signals(self) -> None:
        self.signal_log.connect(self._append_log)
        self.signal_status.connect(self._set_status)
        self.signal_show_info.connect(self._show_info)
        self.signal_show_error.connect(self._show_error)
        self.signal_ask_yes_no.connect(self._handle_ask_yes_no)
        self.signal_ask_ok_cancel.connect(self._handle_ask_ok_cancel)

    def _bind_shortcuts(self) -> None:
        QShortcut(QKeySequence("Alt+S"), self, activated=self.start_update)
        QShortcut(QKeySequence("Alt+Z"), self, activated=self.request_close)
        QShortcut(QKeySequence("Alt+L"), self, activated=lambda: self.log_box.setFocus())
        QShortcut(QKeySequence("Alt+T"), self, activated=self._focus_timeout)
        QShortcut(QKeySequence("F1"), self, activated=self.show_help)

    def _announce_startup(self) -> None:
        self.raise_()
        self.activateWindow()
        self.start_btn.setFocus()
        QMessageBox.information(
            self,
            "Aplikacja gotowa",
            "OpenGD77 Updater zostal uruchomiony.\n"
            "Uzyj Tab aby przechodzic po kontrolkach.\n"
            "Nacisnij Alt+S, aby rozpoczac.",
        )

    def log(self, msg: str) -> None:
        self.signal_log.emit(msg)

    @Slot(str)
    def _append_log(self, msg: str) -> None:
        self.log_box.appendPlainText(msg)
        self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())
        if msg.strip():
            self.last_msg.setText(msg.strip())

    @Slot(str)
    def _set_status(self, text: str) -> None:
        self.status.setText(text)
        self.last_msg.setText(text)

    @Slot(str, str)
    def _show_info(self, title: str, text: str) -> None:
        QMessageBox.information(self, title, text)

    @Slot(str, str)
    def _show_error(self, title: str, text: str) -> None:
        QMessageBox.critical(self, title, text)

    def _focus_timeout(self) -> None:
        self.timeout.setFocus()
        self.timeout.selectAll()

    def show_help(self) -> None:
        QMessageBox.information(
            self,
            "Pomoc",
            "Sterowanie:\n"
            "- Alt+S: Start aktualizacji\n"
            "- Alt+Z: Zamknij\n"
            "- Alt+F4: Zamknij (z potwierdzeniem)\n"
            "- Alt+L: Przejdz do logu\n"
            "- Alt+T: Przejdz do timeout\n"
            "- F1: Ta pomoc",
        )

    def _set_busy(self, busy: bool, status_text: str) -> None:
        self.is_busy = busy
        self.signal_status.emit(status_text)
        self.start_btn.setEnabled(not busy)
        self.close_btn.setEnabled(not busy)
        self.timeout.setEnabled(not busy)
        self.check_only.setEnabled(not busy)
        self.force.setEnabled(not busy)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self.is_busy or (self.worker and self.worker.is_alive()):
            QMessageBox.warning(self, "Aktualizacja w toku", "Nie mozna zamknac programu podczas aktualizacji.")
            event.ignore()
            return
        confirm = QMessageBox.question(
            self,
            "Potwierdzenie zamkniecia",
            "Czy na pewno chcesz zamknac program?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def request_close(self) -> None:
        self.close()

    @Slot(object)
    def _handle_ask_yes_no(self, req: PromptRequest) -> None:
        buttons = QMessageBox.Yes | QMessageBox.No
        default = QMessageBox.No if req.default_no else QMessageBox.Yes
        ret = QMessageBox.question(self, req.title, req.text, buttons, default)
        req.result = ret == QMessageBox.Yes
        req.event.set()

    @Slot(object)
    def _handle_ask_ok_cancel(self, req: PromptRequest) -> None:
        buttons = QMessageBox.Ok | QMessageBox.Cancel
        ret = QMessageBox.question(self, req.title, req.text, buttons, QMessageBox.Ok)
        req.result = ret == QMessageBox.Ok
        req.event.set()

    def ask_yes_no(self, title: str, text: str, default_no: bool = True) -> bool:
        req = PromptRequest(title=title, text=text, default_no=default_no, event=threading.Event(), result=False)
        self.signal_ask_yes_no.emit(req)
        req.event.wait()
        return req.result

    def ask_ok_cancel(self, title: str, text: str) -> bool:
        req = PromptRequest(title=title, text=text, default_no=True, event=threading.Event(), result=False)
        self.signal_ask_ok_cancel.emit(req)
        req.event.wait()
        return req.result

    def start_update(self) -> None:
        if self.is_busy:
            return
        if core is None:
            self.signal_show_error.emit("Blad startu", f"Nie mozna zaladowac modulu aktualizatora.\n{CORE_IMPORT_ERROR}")
            return

        self._set_busy(True, "Trwa sprawdzanie...")
        self.log("")
        self.log("=== OpenGD77 Auto Updater (Qt GUI) ===")
        check_only = self.check_only.isChecked()
        force = self.force.isChecked()
        timeout_s = int(self.timeout.value())

        self.worker = threading.Thread(
            target=self._run_worker,
            args=(check_only, force, timeout_s),
            daemon=True,
        )
        self.worker.start()

    def _run_worker(self, check_only: bool, force: bool, timeout_s: int) -> None:
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
                up_to_date = current_info["build_datetime"][:8] >= latest_date

            do_flash = not check_only
            if up_to_date and not force:
                self.log("Radio wyglada na aktualne wedlug daty buildu.")
                if check_only:
                    do_flash = False
                else:
                    do_flash = self.ask_yes_no("OpenGD77", "Radio jest aktualne.\nWymusic ponowne flashowanie?")
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

            ready = self.ask_ok_cancel(
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
                if driver_inf and self.ask_yes_no(
                    "Brak DFU", "Nie wykryto DFU.\nSprobowac automatycznie zainstalowac sterownik WinUSB?"
                ):
                    self.log(f"Instalacja sterownika z: {driver_inf}")
                    ok, out = core.install_winusb_driver(driver_inf, elevate=True)
                    if out:
                        self.log(out)
                    if ok:
                        retry = self.ask_ok_cancel(
                            "DFU",
                            "Sterownik zainstalowany.\n"
                            "Ustaw ponownie radio w DFU (PTT + S1) i kliknij OK, aby kontynuowac.",
                        )
                        if not retry:
                            self.log("Przerwano po instalacji sterownika.")
                            exit_code = 0
                            return
                        self.log("Ponowne oczekiwanie na DFU...")
                        if not core.wait_for_dfu(libusb_dll, timeout_s=timeout_s):
                            self.signal_show_error.emit("Brak DFU", "Dalej nie wykryto DFU po instalacji sterownika.")
                            exit_code = 4
                            return
                    else:
                        self.signal_show_error.emit(
                            "Blad sterownika", "Nie udalo sie automatycznie zainstalowac sterownika WinUSB."
                        )
                        exit_code = 4
                        return
                else:
                    self.signal_show_error.emit("Brak DFU", "Nie wykryto DFU. Sprawdz radio i sterownik WinUSB.")
                    exit_code = 4
                    return

            self.log("DFU wykryte.")
            if not self.ask_yes_no("OpenGD77", "Rozpoczac flashowanie teraz?"):
                self.log("Przerwano przed flashowaniem.")
                exit_code = 0
                return

            self.log("Start flashowania...")
            rc = core.run_loader_with_backend(loader_path, libusb_dll, ["-s", str(donor_bin), "-m", "MD-UV380", "-f", str(firmware_bin)])
            if rc != 0:
                self.log(f"Flash zakonczyl sie bledem, kod={rc}")
                self.signal_show_error.emit("Blad flashowania", f"Flash zakonczyl sie bledem (kod {rc}).")
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
            self.signal_show_info.emit("Sukces", "Aktualizacja zakonczona.")

        except Exception as ex:
            self.log(f"Blad: {ex}")
            self.log(traceback.format_exc())
            self.signal_show_error.emit("Blad", f"Wystapil blad:\n{ex}")
            exit_code = 10
        finally:
            final = "Zakonczono" if exit_code == 0 else f"Blad (kod {exit_code})"
            QTimer.singleShot(0, lambda: self._set_busy(False, final))

    def _flush_logs(self) -> None:
        # Kept for parity if queue logging is used in future.
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._append_log(msg)
        except queue.Empty:
            pass


def main() -> int:
    app = QApplication(sys.argv)
    if core is None:
        QMessageBox.critical(None, "Blad startu", f"Nie mozna uruchomic OpenGD77 Updater.\n{CORE_IMPORT_ERROR}")
        return 2
    win = UpdaterWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
