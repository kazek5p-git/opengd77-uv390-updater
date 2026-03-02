Aktualizator programu OpenGD77 UV390 (A11y)
==========================================

1) Konfiguracja klienta
- Plik: program_update_config.json
- Pola:
  - app_version: lokalna wersja programu
  - manifest_url: URL do latest.json na serwerze
  - auto_check_on_start: true/false, automatyczny check update przy starcie

2) Build paczki update (lokalnie)
- Polecenie:
  powershell -ExecutionPolicy Bypass -File .\build_program_update_release.ps1 -Version 2026.03.01.2 -UpdateLocalConfig

- Wyniki:
  - ZIP: dist\program_updater_release\<version>\OpenGD77_UV390_A11y_<version>.zip
  - Manifest: dist\program_updater_release\<version>\latest.json

3) Publikacja na kazpar.pl
- Polecenie:
  powershell -ExecutionPolicy Bypass -File .\publish_program_update_release.ps1 -Version 2026.03.01.2 -UpdateLocalConfig

- Domyslnie:
  - host: kazpar.pl
  - port SSH: 1024
  - user: root
  - remote dir: /home/kazek/www/opengd77-updater
  - base URL: https://kazpar.pl/opengd77-updater

4) Uzycie w aplikacji
- W GUI: przycisk "Aktualizuj program" (Alt+A)
- Auto-check przy starcie jest wlaczony, gdy auto_check_on_start=true
- Aplikacja:
  - pobiera manifest,
  - porownuje wersje,
  - pobiera ZIP,
  - weryfikuje SHA256 (jesli jest w manifescie),
  - uruchamia helper podmiany plikow,
  - zamyka i uruchamia sie ponownie.
