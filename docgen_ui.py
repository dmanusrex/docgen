# DocGen - https://github.com/dmanusrex/docgen
#
# Copyright (C) 2023 - Darren Richer
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


''' DocGen Main Screen '''

import os
import pandas as pd
import logging
import customtkinter as ctk
import keyring
import tkinter as tk
from tkinter import filedialog, ttk, BooleanVar, StringVar,  HORIZONTAL
from typing import Any
from tooltip import ToolTip

# Appliction Specific Imports
from config import docgenConfig
from version import DOCGEN_VERSION
from docgen_core import TextHandler, Data_Loader, Generate_Reports, Email_Reports

tkContainer = Any

class _Generate_Documents_Tab(ctk.CTkFrame):   # pylint: disable=too-many-ancestors
    '''Generate Word Documents from a supplied RTR file'''
    def __init__(self, container: tkContainer, config: docgenConfig):
        super().__init__(container)
        self._config = config

        self.df = pd.DataFrame()
        self._officials_list = StringVar(value=self._config.get_str("officials_list"))
        self._report_directory = StringVar(value=self._config.get_str("report_directory"))
        self._report_file = StringVar(value=self._config.get_str("report_file_docx"))
        self._ctk_theme = StringVar(value=self._config.get_str("Theme"))
        self._ctk_size = StringVar(value=self._config.get_str("Scaling"))
        self._ctk_colour = StringVar(value=self._config.get_str("Colour"))
        self._incl_inv_pending = BooleanVar(value=self._config.get_bool("incl_inv_pending"))
        self._incl_pso_pending = BooleanVar(value=self._config.get_bool("incl_pso_pending"))
        self._incl_account_pending = BooleanVar(value=self._config.get_bool("incl_account_pending"))

         # self is a vertical container that will contain 3 frames
        self.columnconfigure(0, weight=1)
        filesframe = ctk.CTkFrame(self)
        filesframe.grid(column=0, row=0, sticky="news")
        filesframe.rowconfigure(0, weight=1)
        filesframe.rowconfigure(1, weight=1)
        filesframe.rowconfigure(2, weight=1)

        optionsframe = ctk.CTkFrame(self)
        optionsframe.grid(column=0, row=2, sticky="news")

        buttonsframe = ctk.CTkFrame(self)
        buttonsframe.grid(column=0, row=4, sticky="news")
        buttonsframe.rowconfigure(0, weight=0)

        # Files Section
        ctk.CTkLabel(filesframe,
            text="Files and Directories").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        btn1 = ctk.CTkButton(filesframe, text="RTR List", command=self._handle_officials_browse)
        btn1.grid(column=0, row=1, padx=20, pady=10)
        ToolTip(btn1, text="Select the RTR officials export file")   # pylint: disable=C0330
        ctk.CTkLabel(filesframe, textvariable=self._officials_list).grid(column=1, row=1, sticky="w")

        btn2 = ctk.CTkButton(filesframe, text="Report Directory", command=self._handle_report_dir_browse)
        btn2.grid(column=0, row=2, padx=20, pady=10)
        ToolTip(btn2, text="Select where output files will be sent")   # pylint: disable=C0330
        ctk.CTkLabel(filesframe, textvariable=self._report_directory).grid(column=1, row=2, sticky="w")

        btn3 = ctk.CTkButton(filesframe, text="Master Report File Name", command=self._handle_report_file_browse)
        btn3.grid(column=0, row=3, padx=20, pady=10)
        ToolTip(btn3, text="Set report file name")   # pylint: disable=C0330
        ctk.CTkLabel(filesframe, textvariable=self._report_file).grid(column=1, row=3, sticky="w")

        # Options Frame - Left and Right Panels

        left_optionsframe = ctk.CTkFrame(optionsframe)
        left_optionsframe.grid(column=0, row=0, sticky="news", padx=10, pady=10)
        left_optionsframe.rowconfigure(0, weight=1)
        right_optionsframe = ctk.CTkFrame(optionsframe)
        right_optionsframe.grid(column=1, row=0, sticky="news", padx=10, pady=10)
        right_optionsframe.rowconfigure(0, weight=1)

        # Program Options on the left frame

        ctk.CTkLabel(left_optionsframe,
            text="UI Appearance").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        ctk.CTkLabel(left_optionsframe, text="Appearance Mode", anchor="w").grid(row=1, column=1, sticky="w")
        ctk.CTkOptionMenu(left_optionsframe, values=["Light", "Dark", "System"],
           command=self.change_appearance_mode_event, variable=self._ctk_theme).grid(row=1, column=0, padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkLabel(left_optionsframe, text="UI Scaling", anchor="w").grid(row=2, column=1, sticky="w")
        ctk.CTkOptionMenu(left_optionsframe, values=["80%", "90%", "100%", "110%", "120%"],
           command=self.change_scaling_event, variable=self._ctk_size).grid(row=2, column=0, padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkLabel(left_optionsframe, text="Colour (Application Restart Required)", anchor="w").grid(row=3, column=1, sticky="w")
        ctk.CTkOptionMenu(left_optionsframe, values=["blue", "green", "dark-blue"],
           command=self.change_colour_event, variable=self._ctk_colour).grid(row=3, column=0, padx=20, pady=10) # pylint: disable=C0330


        # Right options frame for status options
        ctk.CTkLabel(right_optionsframe,
            text="RTR Officials Status").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        ctk.CTkSwitch(right_optionsframe, text = "PSO Pending", variable=self._incl_pso_pending, onvalue = True, offvalue=False,
            command=self._handle_incl_pso_pending).grid(column=0, row=1, sticky="w", padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkSwitch(right_optionsframe, text = "Account Pending", variable=self._incl_account_pending, onvalue = True, offvalue=False,
            command=self._handle_incl_account_pending).grid(column=0, row=2, sticky="w", padx=20, pady=10) # pylint: disable=C0330

        ctk.CTkSwitch(right_optionsframe, text = "Invoice Pending", variable=self._incl_inv_pending, onvalue = True, offvalue=False,
               command=self._handle_incl_inv_pending).grid(column=0, row=3, sticky="w", padx=20, pady=10) # pylint: disable=C0330

        # Add Command Buttons

        ctk.CTkLabel(buttonsframe,
            text="Actions").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        self.load_btn = ctk.CTkButton(buttonsframe, text="Load Datafile", command=self._handle_load_btn)
        self.load_btn.grid(column=0, row=1, sticky="news", padx=20, pady=10)
        self.reset_btn = ctk.CTkButton(buttonsframe, text="Reset", command=self._handle_reset_btn)
        self.reset_btn.grid(column=1, row=1, sticky="news", padx=20, pady=10)
        self.reports_btn = ctk.CTkButton(buttonsframe, text="Generate Reports", command=self._handle_reports_btn)
        self.reports_btn.grid(column=2, row=1, sticky="news", padx=20, pady=10)

    def _handle_officials_browse(self) -> None:
        directory = filedialog.askopenfilename()
        if len(directory) == 0:
            return
        self._config.set_str("officials_list", directory)
        self._officials_list.set(directory)

    def _handle_report_dir_browse(self) -> None:
        directory = filedialog.askdirectory()
        if len(directory) == 0:
            return
        directory = os.path.normpath(directory)
        self._config.set_str("report_directory", directory)
        self._report_directory.set(directory)

    def _handle_report_file_browse(self) -> None:
        report_file = filedialog.asksaveasfilename( filetypes = [('Word Documents','*.docx')], defaultextension=".docx", title="Report File", 
                                                initialfile=os.path.basename(self._report_file.get()), # pylint: disable=C0330
                                                initialdir=self._config.get_str("report_directory")) # pylint: disable=C0330
        if len(report_file) == 0:
            return
        self._config.set_str("report_file_docx", report_file)
        self._report_file.set(report_file)

    def _handle_incl_pso_pending(self, *_arg) -> None:
        self._config.set_bool("incl_pso_pending", self._incl_pso_pending.get())

    def _handle_incl_account_pending(self, *_arg) -> None:
        self._config.set_bool("incl_account_pending", self._incl_account_pending.get())

    def _handle_incl_inv_pending(self, *_arg) -> None:
        self._config.set_bool("incl_inv_pending", self._incl_inv_pending.get())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)
        self._config.set_str("Theme", new_appearance_mode)

    def change_scaling_event(self, new_scaling: str) -> None:
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)
        self._config.set_str("Scaling", new_scaling)

    def change_colour_event(self, new_colour: str) -> None:
        logging.info("Changing colour to : " + new_colour)
        ctk.set_default_color_theme(new_colour)
        self._config.set_str("Colour", new_colour)

    def buttons(self, newstate) -> None:
        '''Enable/disable all buttons'''
        self.load_btn.configure(state = newstate)
        self.reset_btn.configure(state = newstate)
        self.reports_btn.configure(state = newstate)

    def _handle_reports_btn(self) -> None:
        if self.df.empty:
            logging.info ("Load data first...")
            return
        self.buttons("disabled")
        reports_thread = Generate_Reports(self.df, self._config)
        reports_thread.start()
        self.monitor_reports_thread(reports_thread)


    def _handle_load_btn(self) -> None:
        self.buttons("disabled")
        load_thread = Data_Loader(self._config)
        load_thread.start()
        self.monitor_load_thread(load_thread)

    def _handle_reset_btn(self) -> None:
        self.buttons("disabled")
        self.df = pd.DataFrame()
        logging.info("Reset Complete")
        self.buttons("enabled")


    def monitor_load_thread(self, thread):
        if thread.is_alive():
            # check the thread every 100ms 
            self.after(100, lambda: self.monitor_load_thread(thread))
        else:
            # Retrieve data from the loading process and merge it with already loaded data
            if self.df.empty:
                self.df = thread.df
            else:
                self.df = pd.concat([self.df,thread.df], axis=0).drop_duplicates()
                logging.info("%d officials records merged" % self.df.shape[0])
 
            self.buttons("enabled")
            thread.join()

    def monitor_reports_thread(self, thread):
        if thread.is_alive():
            # check the thread every 100ms 
            self.after(100, lambda: self.monitor_reports_thread(thread))
        else:
            self.buttons("enabled")
            thread.join()

class _Email_Documents_Tab(ctk.CTkFrame):   # pylint: disable=too-many-ancestors
    '''E-Mail Completed list of Word Documents'''
    def __init__(self, container: tkContainer, config: docgenConfig):
        super().__init__(container)
        self._config = config

        self._report_directory = StringVar(value=self._config.get_str("report_directory"))
        self._email_list_csv = StringVar(value=self._config.get_str("email_list_csv"))
        self._email_smtp_server = StringVar(value=self._config.get_str("email_smtp_server"))
        self._email_smtp_port = StringVar(value=self._config.get_str("email_smtp_port"))
        self._email_smtp_user = StringVar(value=self._config.get_str("email_smtp_user"))
        self._email_from = StringVar(value=self._config.get_str("email_from"))
        self._email_subject = StringVar(value=self._config.get_str("email_subject"))
        self._email_body = self._config.get_str("email_body")

         # self is a vertical container that will contain 3 frames
        self.columnconfigure(0, weight=1)

        filesframe = ctk.CTkFrame(self)
        filesframe.grid(column=0, row=0, sticky="news")
        filesframe.rowconfigure(0, weight=1)
        filesframe.rowconfigure(1, weight=1)

        optionsframe = ctk.CTkFrame(self)
        optionsframe.grid(column=0, row=2, sticky="news")

        buttonsframe = ctk.CTkFrame(self)
        buttonsframe.grid(column=0, row=4, sticky="news")
        buttonsframe.rowconfigure(0, weight=0)

        # Files Section
        ctk.CTkLabel(filesframe,
            text="E-mail Configuration").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        # options Section


        entry_width = 600

#        reg_email_smtp_server = self.register(self._handle_email_smtp_server)
#       A registered validation function seems to disable the interactive logging window. Need to investigate

        ctk.CTkLabel(optionsframe, text="SMTP Server", anchor="w").grid(row=1, column=0, sticky="w")

        smtp_server_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_smtp_server, width=entry_width)
        smtp_server_entry.grid(column=1, row=1, sticky="w", padx=10, pady=10)
        smtp_server_entry.bind('<FocusOut>', self._handle_email_smtp_server)

        ctk.CTkLabel(optionsframe, text="SMTP Port", anchor="w").grid(row=2, column=0, sticky="w")
        smtp_port_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_smtp_port, width=entry_width)
        smtp_port_entry.grid(column=1, row=2, sticky="w", padx=10, pady=10)
        smtp_port_entry.bind('<FocusOut>', self._handle_email_smtp_port)
        
        ctk.CTkLabel(optionsframe, text="SMTP Username", anchor="w").grid(row=3, column=0, sticky="w")
        smtp_user_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_smtp_user, width=entry_width)
        smtp_user_entry.grid(column=1, row=3, sticky="w", padx=10, pady=10)
        smtp_user_entry.bind('<FocusOut>', self._handle_email_smtp_user)

        ctk.CTkLabel(optionsframe, text="SMTP Password", anchor="w").grid(row=4, column=0, sticky="w")
        password_entry = ctk.CTkEntry(optionsframe, placeholder_text="Password", show="*", width=entry_width)
        password_entry.grid(column=1, row=4, sticky="w", padx=10, pady=10)
        password_entry.bind('<FocusOut>', self._handle_email_smtp_password)

        ctk.CTkLabel(optionsframe, text="E-mail From", anchor="w").grid(row=5, column=0, sticky="w")
        email_from_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_from, width=entry_width)
        email_from_entry.grid(column=1, row=5, sticky="w", padx=10, pady=10)
        email_from_entry.bind('<FocusOut>', self._handle_email_from)

        ctk.CTkLabel(optionsframe, text="E-mail Subject", anchor="w").grid(row=6, column=0, sticky="w")
        email_subject_entry = ctk.CTkEntry(optionsframe, textvariable=self._email_subject, width=entry_width)
        email_subject_entry.grid(column=1, row=6, sticky="w", padx=10, pady=10)
        email_subject_entry.bind('<FocusOut>', self._handle_email_subject)

        # Body Text
        ctk.CTkLabel(optionsframe, text="E-mail Body", anchor="w").grid(row=7, column=0, sticky="w")

        self.txtbodybox = ctk.CTkTextbox(master=optionsframe, state='normal', width=entry_width)
        self.txtbodybox.grid(column=1, row=7, sticky="w", padx=10, pady=10)
        self.txtbodybox.insert(tk.END, self._email_body)

        # Add Command Buttons

        ctk.CTkLabel(buttonsframe,
            text="Actions").grid(column=0, row=0, sticky="w", padx=10)   # pylint: disable=C0330

        self.emailtest_btn = ctk.CTkButton(buttonsframe, text="Send Test EMails", command=self._handle_email_test_btn)
        self.emailtest_btn.grid(column=0, row=1, sticky="news", padx=20, pady=10)
        self.emailall_btn = ctk.CTkButton(buttonsframe, text="Send All Emails", command=self._handle_email_all_btn)
        self.emailall_btn.grid(column=1, row=1, sticky="news", padx=20, pady=10)


    def _handle_report_dir_browse(self) -> None:
        directory = filedialog.askdirectory()
        if len(directory) == 0:
            return
        directory = os.path.normpath(directory)
        self._config.set_str("report_directory", directory)
        self._report_directory.set(directory)

    def _handle_email_smtp_server(self, event) -> bool:
        self._config.set_str("email_smtp_server", event.widget.get())
        return True

    def _handle_email_smtp_port(self, event) -> bool:
        self._config.set_str("email_smtp_port", event.widget.get())
        return True
    
    def _handle_email_smtp_user(self, event) -> bool:
        self._config.set_str("email_smtp_user", event.widget.get())
        return True

    def _handle_email_smtp_password(self, event) -> bool:
        if event.widget.get() != "Password":
            keyring.set_password("SWON-DOCGEN", self._email_smtp_user.get(), event.widget.get())
            logging.info("Password Changed for %s" % self._email_smtp_user.get())
        return True
    
    def _handle_email_from(self, event) -> bool:
        self._config.set_str("email_from", event.widget.get())
        return True
    
    def _handle_email_subject(self, event) -> bool:
        self._config.set_str("email_subject", event.widget.get())
        return True
    
    def _handle_email_body(self, event) -> bool:
        self._config.set_str("email_body", event.widget.get())
        return True



    def buttons(self, newstate) -> None:
        '''Enable/disable all buttons'''
        self.emailtest_btn.configure(state = newstate)
        self.emailall_btn.configure(state = newstate)

    def _handle_email_test_btn(self) -> None:
#        if self.df.empty:
#            logging.info ("Load data first...")
#            return
        self.buttons("disabled")
        email_thread = Email_Reports(True, self._config)
        email_thread.start()
        self.monitor_email_thread(email_thread)

    def _handle_email_all_btn(self) -> None:
#        if self.df.empty:
#            logging.info ("Load data first...")
#            return
        self.buttons("disabled")
        email_thread = Email_Reports(False, self._config)
        email_thread.start()
        self.monitor_email_thread(email_thread)

    def monitor_email_thread(self, thread):
        if thread.is_alive():
            # check the thread every 100ms 
            self.after(100, lambda: self.monitor_email_thread(thread))
        else:
            self.buttons("enabled")
            thread.join()


class _Logging(ctk.CTkFrame): # pylint: disable=too-many-ancestors,too-many-instance-attributes
    '''Logging Window'''
    def __init__(self, container: ctk.CTk, config: docgenConfig):
        super().__init__(container) 
        self._config = config
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Messages").grid(column=0, row=0, sticky="ws", pady=10)

        self.logwin = ctk.CTkTextbox(self, state='disabled')
        self.logwin.grid(column=0, row=4, sticky='nsew')
        # Logging configuration
        logfile = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "docgen.log"))

        logging.basicConfig(filename=logfile,
                            level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        # Create textLogger
        text_handler = TextHandler(self.logwin)
        text_handler.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))
        # Add the handler to logger
        logger = logging.getLogger()
        logger.addHandler(text_handler)

class docgenApp(ctk.CTkFrame):  # pylint: disable=too-many-ancestors
    '''Main Appliction'''
    # pylint: disable=too-many-arguments,too-many-locals
    def __init__(self, container: ctk.CTk,
                 config: docgenConfig):
        super().__init__(container)
        self._config = config

        self.grid(column=0, row=0, sticky="news")
        self.columnconfigure(0, weight=1)
        # Odd rows are empty filler to distribute vertical whitespace
        for i in [1, 3]:
            self.rowconfigure(i, weight=1)

        self.tabview = ctk.CTkTabview(self, width=container.winfo_width())
        self.tabview.grid(row=0, column=0, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.tabview.add("Generate Documents")
        self.tabview.add("E-Mail Documents")

        # Generate Documents Tab
        self.tabview.tab("Generate Documents").grid_columnconfigure(0, weight=1)
        self.configfiles = _Generate_Documents_Tab(self.tabview.tab("Generate Documents"), self._config)
        self.configfiles.grid(column=0, row=0, sticky="news")


        # E-Mail Documents Tab
        self.tabview.tab("E-Mail Documents").grid_columnconfigure(0, weight=1)
        self.emailtab = _Email_Documents_Tab(self.tabview.tab("E-Mail Documents"), self._config)
        self.emailtab.grid(column=0, row=0, sticky="news")

        # Logging Window
        loggingwin = _Logging(self, self._config)
        loggingwin.grid(column=0, row=2, padx=(20, 20), pady=(20, 0), sticky="news")
 
        # Info panel
        fr8 = ctk.CTkFrame(self)
        fr8.grid(column=0, row=4, sticky="news")
        fr8.rowconfigure(0, weight=1)
        fr8.columnconfigure(0, weight=1)
        link_label = ctk.CTkLabel(fr8,
            text="Documentation: Maybe someday!")  # pylint: disable=C0330
        link_label.grid(column=0, row=0, sticky="w")
        version_label = ctk.CTkLabel(fr8, text="Version "+DOCGEN_VERSION)
        version_label.grid(column=1, row=0, sticky="nes")


        
def main():
    '''testing'''
    root = ctk.CTk()
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    root.resizable(True, True)
    options = docgenConfig()
    settings = docgenApp(root, options)
    settings.grid(column=0, row=0, sticky="news")
    logging.info("Hello World")
    root.mainloop()

if __name__ == '__main__':
    main()
