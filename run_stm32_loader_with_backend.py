import runpy
import sys

import usb.core
import usb.backend.libusb1

LOADER = r"C:\Users\Kazek\OpenGD77_UV390_Plus10W\stm32_loader_tools\OpenGD77_MDUV380_DM1701_20260130\MDUV380_firmware\tools\opengd77_stm32_firmware_loader.py"
LIBUSB_DLL = r"C:\Users\Kazek\OpenGD77_UV390_Plus10W\libusb\dll64\MinGW64\dll\libusb-1.0.dll"

backend = usb.backend.libusb1.get_backend(find_library=lambda name: LIBUSB_DLL)
if backend is None:
    print("ERROR: Unable to initialize libusb backend")
    sys.exit(2)

orig_find = usb.core.find

def patched_find(*args, **kwargs):
    kwargs.setdefault("backend", backend)
    return orig_find(*args, **kwargs)

usb.core.find = patched_find

sys.argv = [LOADER] + sys.argv[1:]
runpy.run_path(LOADER, run_name="__main__")
