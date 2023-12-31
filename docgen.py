# Officials Doc Gen - https://github.com/dmanusrex/docgen
#
# Copyright (C) 2021 - Darren Richer
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
# IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
# OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
# OR OTHER DEALINGS IN THE SOFTWARE.

'''Produce a report of the officials in a club and recommendations on activiaties'''

import customtkinter as ctk
import docgen_ui as ui
from config import docgenConfig
import docgen_version
import os
import sys
import logging
from requests.exceptions import RequestException

from version import DOCGEN_VERSION


def check_for_update() -> None:
    """Notifies if there's a newer released version"""
    current_version = DOCGEN_VERSION
    try:
        latest_version = docgen_version.latest()
        if latest_version is not None and not docgen_version.is_latest_version(
            latest_version, current_version
        ):
            logging.info(
                f"New version available {latest_version.tag}"
            )
            logging.info(f"Download URL: {latest_version.url}")
#           Make it clickable???  webbrowser.open(latest_version.url))
    except RequestException as ex:
        logging.warning("Error checking for update: %s", ex)


def main():
    '''Runs the application'''

    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))

    root = ctk.CTk()
    config = docgenConfig()
    ctk.set_appearance_mode(config.get_str("Theme"))  # Modes: "System" (standard), "Dark", "Light"
    ctk.set_default_color_theme(config.get_str("Colour"))  # Themes: "blue" (standard), "green", "dark-blue"
    new_scaling_float = int(config.get_str("Scaling").replace("%", "")) / 100
#    ctk.set_widget_scaling(new_scaling_float)
#    ctk.set_window_scaling(new_scaling_float)
    root.title("Swim Ontario - Officials Doc Generator")
    icon_file = os.path.abspath(os.path.join(bundle_dir, 'media','swon-logo.ico'))
    root.iconbitmap(icon_file)
#    root.geometry(f"{850}x{10500}")
#    root.minsize(800, 900)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.resizable(True, True)
    content = ui.docgenApp(root, config)
    content.grid(column=0, row=0, sticky="news")
    check_for_update()

    try:
        root.update()
        # pylint: disable=import-error,import-outside-toplevel
        import pyi_splash  # type: ignore

        if pyi_splash.is_alive():
            pyi_splash.close()
    except ModuleNotFoundError:
        pass
    except RuntimeError:
        pass

    root.mainloop()

    config.save()

if __name__ == "__main__":
    main()

