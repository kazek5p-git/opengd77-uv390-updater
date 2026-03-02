OpenGD77 UV390 Plus 10W - uruchamianie GUI

Wersje ponizej nie wymagaja instalacji Pythona ani dodatkowych komponentow.
Interfejs ma poprawki pod czytnik ekranu NVDA (status, ostatni komunikat, skroty klawiaturowe i komunikat startowy).

1) Najprosciej:
   - kliknij skrot na pulpicie:
     OpenGD77 UV390 Updater.lnk
   - Skroty: Alt+S Sprawdz, Alt+T Start, Alt+A Aktualizuj program, Alt+L Log, Alt+D Timeout, Alt+F4 Zamknij, F1 Pomoc.
   - Tab / Shift+Tab: przechodzenie po kontrolkach.
   - Alt+F4 i przycisk X pytaja o potwierdzenie zamkniecia.
   - W trakcie aktualizacji zamkniecie programu jest blokowane.
   - Aktualizator programu pobiera manifest z:
     https://kazpar.pl/opengd77-updater/latest.json
   - Auto-check aktualizacji programu przy starcie: wlaczony (konfigurowalne w program_update_config.json).
   - Log diagnostyczny crashy:
     C:\Users\Kazek\OpenGD77_UV390_Plus10W\logs\OpenGD77_A11y_YYYYMMDD_HHMMSS.log

2) Wersja A11y (natywne kontrolki Windows Forms):
   - Skrypt GUI:
     C:\Users\Kazek\OpenGD77_UV390_Plus10W\OpenGD77_UV390_A11y.ps1
   - Launcher bez konsoli:
     C:\Users\Kazek\OpenGD77_UV390_Plus10W\start_OpenGD77_UV390_A11y.vbs
   - Backend CLI:
     C:\Users\Kazek\OpenGD77_UV390_Plus10W\dist\OpenGD77_UV390_BackendCLI.exe

3) Wersja przenosna ONE-FILE (jeden plik EXE):
   - C:\Users\Kazek\OpenGD77_UV390_Plus10W\dist\OpenGD77_UV390_Updater_OneFile.exe
   - Paczka ZIP:
     C:\Users\Kazek\OpenGD77_UV390_Plus10W\OpenGD77_UV390_Updater_OneFile_20260301.zip

4) Wersja przenosna FOLDER (EXE + _internal):
   - C:\Users\Kazek\OpenGD77_UV390_Plus10W\dist\OpenGD77_UV390_Updater\OpenGD77_UV390_Updater.exe
   - Paczka ZIP:
     C:\Users\Kazek\OpenGD77_UV390_Plus10W\OpenGD77_UV390_Updater_portable_full_20260301.zip

5) Stare launchery skryptowe (opcjonalnie):
   - C:\Users\Kazek\OpenGD77_UV390_Plus10W\start_auto_update_uv390_plus_gui.vbs
   - C:\Users\Kazek\OpenGD77_UV390_Plus10W\start_auto_update_uv390_plus_gui.ps1
