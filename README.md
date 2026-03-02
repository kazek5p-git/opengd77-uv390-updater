# OpenGD77 UV390 Updater (source)

Repozytorium zawiera zrodla i skrypty pakujace dla updatera OpenGD77 UV390.

## Automatyczne wydania (GitHub Releases)

Release tworzy sie automatycznie po wypchnieciu taga w formacie `vX.Y.Z`.

Przyklad:

```bash
git tag v2026.03.02.1
git push origin main --tags
```

Workflow:

1. Buduje `OpenGD77_UV390_BackendCLI.exe` (PyInstaller).
2. Tworzy paczke updatera `OpenGD77_UV390_A11y_<wersja>.zip`.
3. Tworzy `latest.json` i `*.sha256`.
4. Publikuje pliki jako GitHub Release dla danego taga.
