# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['opengd77_auto_update_uv390_plus.py'],
    pathex=[],
    binaries=[],
    datas=[('libusb/dll64/MinGW64/dll/libusb-1.0.dll', 'resources'), ('donor_extracted/MD9600-CSV(2571V5)-V26.45.bin', 'resources'), ('wdi_driver_extract/usb_device.inf', 'resources/driver'), ('wdi_driver_extract/usb_device.cat', 'resources/driver'), ('stm32_loader_tools/OpenGD77_MDUV380_DM1701_20260130/MDUV380_firmware/tools/opengd77_stm32_firmware_loader.py', 'resources')],
    hiddenimports=[
        'serial',
        'serial.tools.list_ports',
        'usb',
        'usb.core',
        'usb.util',
        'usb.backend.libusb1',
        # Imported dynamically by bundled STM32 loader executed via runpy.
        'configparser',
        'collections',
        'inspect',
        'zlib',
        'hashlib',
        'enum',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='OpenGD77_UV390_BackendCLI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
