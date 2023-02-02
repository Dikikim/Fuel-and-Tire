import os
import re
import shutil
import subprocess
import threading
import time
import tkinter as tk
from datetime import datetime
from decimal import Decimal
from functools import partial
from io import BytesIO
from shutil import copyfile, SameFileError
from tkinter import ttk, messagebox, filedialog
from typing import Optional

import matplotlib.pyplot as plt
import numpy as np
import paramiko
import requests
from PIL import ImageTk, Image
from uptime import uptime

import aprivatoken
import audiohandler
import backup_files
import bulk_charge
import constants
import email_reciept
import employee_log
import fts_util
import fts_widgets
import genanyreceipt
import main_window
import otireader
import pigraph
from calibration import CalibrationWindow
from fts_util import guarantee_message_send, run_main_thread
from globals import Maint, Data
from logger import logger
from pdfgen import PDFGen
from photocache import PhotoCache
from pimagic import Pi, pyro_run
from popups import COFPopup
from printpdf import convertandprint, convertpdf, printimage
from sqlmanager import Database
from otistructs import Status


class MaintenanceScreen(tk.Frame):
    instance: 'MaintenanceScreen' = None  # holds the main reference to the maintenance screen for others to use

    def __init__(self, master, white=None, **kw):
        super().__init__(master, **kw)
        self.root = master
        self.SX = master.winfo_width()
        self.SY = master.winfo_height()

        maintenance_screen_img = PhotoCache.get("N24_KioskGraphic_17.5x10_Maintenance2018_v3.png")
        tk.Label(self, bd=0, image=maintenance_screen_img).place(x=0, y=0, relheight=1, relwidth=1)

        # the grid
        #         |  0.04175          0.2245           0.4072           0.5900           0.7727
        #         |  X================X================X================X================X================
        # 0.1800  |  .                .                valve test       .                bulk charge
        # 0.2610  |  obc gen          .                .                defrag           coupon
        # 0.3465  |  leak test        .                .                restore inf.     promo
        # 0.4320  |  calib            send email       Flowrate test    kiosk info       clock in
        # 0.5155  |  hp calib         graphing         tire response    valve timer      grapher
        # 0.5975  |  .                .                X================X================X================
        #         |  X================X================X

        maintenance_screen_logo = PhotoCache.get("AASTRAEA logo.png")
        tk.Label(self, bd=0, image=maintenance_screen_logo).place(x=2245, y=1800, relheight=0.0610, relwidth=0.0182)
        txtbar = tk.Text(self, text=Database.select('minions', ('id', 'name')))
        txtbar = {"Minion": ['id'] + ['name'] in txtbar}
        txtbar.configure(highlightthickness=0, bd=0, relief='flat')
        txtbar.place(relx=0.5900, rely=0)
        obcbtn = tk.Button(self, image=PhotoCache.get("maint_obc_btn.png"), command=self.gen_obc)  # done
        obcbtn.configure(highlightthickness=0, bd=0, relief='flat')
        obcbtn.place(relx=0.04175, rely=0.2610)
        sysleak = tk.Button(self, image=PhotoCache.get("systemleaktest.png"), command=self.leak_test_pop)  # done
        sysleak.configure(highlightthickness=0, bd=0, relief='flat')
        sysleak.place(relx=0.04175, rely=0.3465)
        syscaltf = tk.Button(self, image=PhotoCache.get("15ptsystemcalibration.png"), command=self.calib_pop)  # done
        syscaltf.configure(highlightthickness=0, bd=0, relief='flat')
        syscaltf.place(relx=0.04175, rely=0.4320)
        hpcalibbtn = tk.Button(self, image=PhotoCache.get("hpcalib.png"), command=self.hpcalib_pop)  # done
        hpcalibbtn.configure(highlightthickness=0, bd=0, relief='flat')
        hpcalibbtn.place(relx=0.04175, rely=0.5155)
        self.mines = tk.Button(self, image=PhotoCache.get("minionemailstatusandtest.png"))
        self.mines.configure(highlightthickness=0, bd=0, relief='flat')
        self.mines.place(relx=0.2245, rely=0.4310)
        graphbtn = tk.Button(self, image=PhotoCache.get("maint_graphing_button.png"), command=self.graphing_pop)
        graphbtn.configure(highlightthickness=0, bd=0, relief='flat')
        graphbtn.place(relx=0.2245, rely=0.5155)
        graphbtn = tk.Button(self, image=PhotoCache.get("tireresponse_button.png"), command=self.tire_response_pop)
        graphbtn.configure(highlightthickness=0, bd=0, relief='flat')
        graphbtn.place(relx=0.4072, rely=0.5155)
        valvact = tk.Button(self, image=PhotoCache.get("valveactuation.png"), command=self.valve_control)
        valvact.configure(highlightthickness=0, bd=0, relief='flat')
        valvact.place(relx=0.40725, rely=0.1800)
        frt = tk.Button(self, image=PhotoCache.get("flowratetest.png"), command=self.flowrate_test)
        frt.configure(highlightthickness=0, bd=0, relief='flat')
        frt.place(relx=0.4072, rely=0.4320)
        driveop = tk.Button(self, image=PhotoCache.get("driveoptimization.png"),
                            command=self.driveOp_pop)  # calls defrag
        driveop.configure(highlightthickness=0, bd=0, relief='flat')
        driveop.place(relx=0.5900, rely=0.2610)
        employee_clock_button = tk.Button(self, image=PhotoCache.get("employee_clock_button.png"),
                                          command=self.employee_clock_popup)
        employee_clock_button.configure(highlightthickness=0, bd=0, relief='flat')
        employee_clock_button.place(relx=0.7727, rely=0.4320)
        kinfo = tk.Button(self, image=PhotoCache.get("kioskinformation.png"), command=self.kiosk_info)
        kinfo.configure(highlightthickness=0, bd=0, relief='flat')
        kinfo.place(relx=0.5900, rely=0.4320)
        filltime = tk.Button(self, image=PhotoCache.get("fillvalvetime.png"), command=self.fill_valve_timing)
        filltime.configure(highlightthickness=0, bd=0, relief='flat')
        filltime.place(relx=0.5900, rely=0.5155)
        # opschedule_btn = tk.Button(self, image=PhotoCache.get("opschedule_img.png"), command=self.operation_schedule_ticket_gui)
        # opschedule_btn.configure(highlightthickness=0, bd=0, relief='flat')
        # opschedule_btn.place(relx=0.04175, rely=0.5975)
        rmg_btn = tk.Button(self, image=PhotoCache.get("maint_coupon_button.png"), command=self.change_coupon)
        rmg_btn.configure(highlightthickness=0, bd=0, relief='flat')
        rmg_btn.place(relx=0.7727, rely=0.2610)
        promo_btn = tk.Button(self, image=PhotoCache.get("maint_promo_button.png"), command=self.change_promo)
        promo_btn.configure(highlightthickness=0, bd=0, relief='flat')
        promo_btn.place(relx=0.7727, rely=0.3465)
        gengraph_button = tk.Button(self, image=PhotoCache.get("maint_gen_graph_button.png"),
                                    command=self.gen_graph_popup)
        gengraph_button.configure(highlightthickness=0, bd=0, relief='flat')
        gengraph_button.place(relx=0.7727, rely=0.5150)
        bulkcharge_button = tk.Button(self, image=PhotoCache.get("maint_bulk_charge_button.png"),
                                      command=self.bulk_charge_popup)
        bulkcharge_button.configure(highlightthickness=0, bd=0, relief='flat')
        bulkcharge_button.place(relx=0.7727, rely=0.1800)
        restore_inf_button = tk.Button(self, image=PhotoCache.get("restore_inf_button.png"),
                                       command=self.restore_inflation_popup)
        restore_inf_button.configure(highlightthickness=0, bd=0, relief='flat')
        restore_inf_button.place(relx=0.5900, rely=0.3465)

        # buttons that we don't use anymore, but could

        # sysclean = Button(maintenance_screen, image = photo1, command = checklist_pop) #done
        # sysclean.configure(highlightthickness=0, bd=0, relief='flat')
        # sysclean.place(relx = 0.04185, rely = 0.1800)
        # syscalsf = Button(maintenance_screen, image = photo5, command = calib_pop) #done
        # syscalsf.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # syscalsf.place(relx = 0.04175, rely = 0.5155)
        # tankcap = Button(maintenance_screen, image = photo6, command = tankcapacitytest)
        # tankcap.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # tankcap.place(relx = 0.2245, rely = 0.1800)
        # printmvt = Button(maintenance_screen, image=photo7, command=print_CALIBRATION_receipt)
        # printmvt.configure(highlightthickness=0, bd=0, relief='flat')
        # printmvt.place(relx=0.2245, rely=0.2630)
        # qumail = Button(maintenance_screen, image = photo10, command = mail_pop)
        # qumail.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # qumail.place(relx = 0.2245, rely = 0.5150)
        # file_maint = Button(maintenance_screen, image = photo23, command = move_maint)
        # file_maint.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # file_maint.place(relx = 0.2245, rely = 0.5990)
        # uploadmvt = Button(maintenance_screen, image = photo12, command = callback)
        # uploadmvt.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # uploadmvt.place(relx = 0.40725, rely = 0.2610)
        # vpntest = Button(maintenance_screen, image = photo13, command = callback)
        # vpntest.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # vpntest.place(relx = 0.40725, rely = 0.3465)
        # onecent = Button(maintenance_screen, image = photo14, command = callback)
        # onecent.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # onecent.place(relx = 0.40725, rely = 0.4320)
        # cctest = Button(maintenance_screen, image = photo15, command = callback)
        # cctest.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # cctest.place(relx = 0.40725, rely = 0.5155)
        # bill_transfer = Button(maintenance_screen, image = photo22, command = move_billing)
        # bill_transfer.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # bill_transfer.place(relx = 0.40725, rely = 0.5990)
        # winup = Button(maintenance_screen, image = photo16, command = windowsUpdate_pop)
        # winup.configure(highlightthickness = 0, bd = 0, relief = 'flat')
        # winup.place(relx = 0.5900, rely = 0.1800)
        # tbd = Button(maintenance_screen, image=photo18, command=tire_bp_cmd)
        # tbd.configure(highlightthickness=0, bd=0, relief='flat')
        # tbd.place(relx=0.5900, rely=0.3465)
        # ctest = Button(maintenance_screen, image = new_cc ,command=Credit_test)
        # ctest.configure(highlightthickness = 0, bd = 0,bg='yellow',fg='green', relief = 'flat')
        # ctest.place(relx=0.7727, rely = 0.1800)

        edit_font = ("Helvetica", "10", "bold")
        enable_btn = tk.Button(self, text="Edit Values", font=edit_font, command=self.enab_maint, bg="#44dd55")
        enable_btn.place(relx=0.08, rely=0.92)
        disable_btn = tk.Button(self, text="Save Values", font=edit_font, command=self.disab_maint, bg="#5555dd")
        disable_btn.place(relx=0.16, rely=0.92)

        self.exmain = tk.Button(self, image=PhotoCache.get("exitmaintenance.png"))
        self.exmain.configure(highlightthickness=0, bd=0, relief='flat')
        self.exmain.place(relx=0.7720, rely=0.9150)

        # Label(maintenance_screen, text = "System Error %:", font=mainnew_small_font).place(relx = 0.848 , rely = 0.62, anchor = E)
        self.create_registered_field("set_presserr", relx=0.1350, rely=0.7175, width=15)
        # Label(maintenance_screen, text = "Hose Press SP %:", font=mainnew_small_font).place(relx = 0.848 , rely = 0.67, anchor = E)
        self.create_registered_field("hose_deflate", relx=0.1370, rely=0.8000, width=15)
        self.create_registered_field("nitrogen_percent", relx=0.4850, rely=0.6340, width=11)
        # Label(maintenance_screen, text = "RMG Setting:", font=mainnew_small_font).place(relx = 0.848 , rely = 0.72, anchor = E)
        self.create_registered_field("rmg_code", relx=0.1240, rely=0.8775, text="(None)", bind=False, width=17,
                                     state=tk.DISABLED)
        # Label(maintenance_screen, text = "Cal:", font=mainnew_small_font).place(relx = 0.698 , rely = 0.77, anchor = E)
        self.create_registered_field("calib_curve", relx=0.6215, rely=0.6370, width=45)
        # Label(maintenance_screen, text = "Inflation Datafile Loc:", font=mainnew_small_font).place(relx = 0.698 , rely = 0.82, anchor = E)
        self.create_registered_field("inf_datafile", relx=0.3500, rely=0.7175, width=84)
        # Label(maintenance_screen, text = "Pmt DF Loc:", font=mainnew_small_font).place(relx = 0.698 , rely = 0.87, anchor = E)
        self.create_registered_field("pay_transfile", relx=0.3485, rely=0.797, width=84)
        # Label(maintenance_screen, text = "Minion Status DF Loc:", font=mainnew_small_font).place(relx = 0.698 , rely = 0.92, anchor = E)
        self.create_registered_field("min_statfile", relx=0.3555, rely=0.8775, width=83)
        # Label(maintenance_screen, text = "Cal DF Loc:", font=mainnew_small_font).place(relx = 0.698 , rely = 0.97, anchor = E)
        self.create_registered_field("max_wait_time", relx=0.3425, rely=0.955, width=8)
        self.create_registered_field("tank_liters", relx=0.4925, rely=0.955, width=9)

        self.disab_maint()

        MaintenanceScreen.instance = self

    def create_registered_field(self, key, relx, rely, text="", bind=True, **kw):
        entry = tk.Entry(self, font=('Helvetica-Bold', 14), bd=10, relief=tk.FLAT, **kw)
        if bind:
            entry.bind('<Return>', self.save_maint)
        entry.place(relx=relx, rely=rely, anchor=tk.W)
        Maint.register_field(key, entry)
        if text:
            Maint.vals[key] = text

    def set_heartbeat_callback(self, func: callable):
        self.mines.config(command=func)

    def set_exit_callback(self, func: callable):
        self.exmain.config(command=func)

    def _config_maint(self, state):
        Maint.vals.config("set_presserr", state=state)
        Maint.vals.config("hose_deflate", state=state)
        Maint.vals.config("calib_curve", state=state)
        Maint.vals.config("inf_datafile", state=state)
        Maint.vals.config("pay_transfile", state=state)
        Maint.vals.config("min_statfile", state=state)
        Maint.vals.config("calib_file", state=state)
        Maint.vals.config("nitrogen_percent", state=state)
        Maint.vals.config("max_wait_time", state=state)
        Maint.vals.config("tank_liters", state=state)

    def enab_maint(self):
        self._config_maint(tk.NORMAL)

    def disab_maint(self):
        self.save_maint()
        self._config_maint(tk.DISABLED)

    def save_maint(self):
        """
		Save entry values to file
		"""
        Maint.vals.refresh()
        try:
            Maint.save_to_json(Maint.config_file)
        except OSError as e:
            logger.error("could not save maintenance values...", exc_info=e)

    def drive_op(self):
        subprocess.call('RUNAS /user:admin "defrag C:"')

    def driveOp_pop(self):
        driveOp_pop_window = tk.Toplevel(self.root)
        driveOp_pop_window.attributes('-topmost', True)
        driveOp_pop_window.title('driveOp_pop_window')
        driveOp_pop_window.update_idletasks()

        w = 300
        h = 100
        x = (self.SX / 2) - (w / 2)
        y = (self.SY / 2) - (h / 2)
        driveOp_pop_window.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.drive_op()
        # t.geometry('128x64+{}+{}'.format(sx/2,sy/2))
        err = tk.Label(driveOp_pop_window, text='Drive Optimization')
        b = tk.Button(driveOp_pop_window, text='O.K.', command=driveOp_pop_window.destroy)
        err.pack()
        b.pack(side=tk.BOTTOM)

    def employee_clock_popup(self):
        # aka technician time clock
        ewin = tk.Toplevel(background="white")
        ewin.attributes('-topmost', 'true')
        ewin.title("Technician Time Clock")
        ewin.geometry('380x200+1000+500')  # 300x220 screen
        locked = False  # this is so it can't run 2 connections at the same time
        first_time = False  # only used if clocking in. used to specify if this should be the start of a new record or not

        def do_select():
            # ran whenever an employee is selected from the box
            # or automatically when the window is open and an employee is already clocked in
            nonlocal locked, first_time
            if locked:
                return
            locked = True
            first_time = False
            employee_name = employee_box.get()
            clock_in_button.place_forget()
            clock_out_button.place_forget()
            send_report_button.place_forget()
            employee_box.config(state=tk.DISABLED)
            etime_var.set('Loading...')
            ewin.update()  # erase all traces of last employee before switching
            msg = ''
            if employee_name != "<Select Employee>":  # is it a valid choice?
                # determine which button to show based on status of clocked-in
                request_params = {"type": "clock-status", "requester": str(Maint.mioskid), "employee": employee_name}
                try:
                    status = requests.get(url=constants.connect_url, params=request_params,
                                          timeout=2.0)  # SENECAWIN7 is Laura's
                    if status.text == "in":  # fully clocked in
                        clock_out_button.place(relx=0.55, rely=0.3)  # if clocked in, allow to clock out
                    elif status.text == "out":  # fully clocked out, no one else clocked in
                        clock_in_button.place(relx=0.13, rely=0.3)  # if clocked out, allow to clock in
                        first_time = True
                    elif status.text == "break":  # clocked out, but could resume or end session
                        send_report_button.place(relx=0.55, rely=0.3)  # if on break, allow to send report
                        clock_in_button.place(relx=0.13, rely=0.3)  # OR clock in again
                    elif status.text.startswith("no"):  # is no
                        employee_in_use = status.text.split("|")[1]
                        lastname = employee_in_use.split(",")[0]
                        msg = "Miosk in use by " + lastname
                    else:  # status.text == "other"
                        miosk_in_use = status.text.split("|")[1]
                        msg = "Already clocked-in on " + miosk_in_use
                except requests.exceptions.RequestException as e:
                    logger.error("Server error", exc_info=e)
                    msg = "Server error."
            locked = False
            etime_var.set(msg)
            employee_box.config(state='readonly')

        def on_select(_event=None):
            threading.Thread(target=do_select, daemon=True, name="employeeselect").start()

        employee_name_map = employee_log.load_active_employees()
        employee_names = list(sorted(employee_name_map.keys()))

        employee_box = ttk.Combobox(ewin, values=["<Select Employee>"] + employee_names,
                                    state='readonly', font=('Helvetica', 14))
        ewin.option_add('*TCombobox*Listbox.font', ('Helvetica', 12))
        employee_box.bind('<<ComboboxSelected>>', on_select)
        employee_box.current(0)  # start on <select>
        employee_box.place(relx=0.01, rely=0.025)

        last_employee_id, last_clockedin = employee_log.load_last_clockedin_details()
        if last_employee_id is not None:
            last_employee_name = employee_log.get_employee_name(last_employee_id)
            if last_employee_name in employee_names:
                employee_box.set(
                    last_employee_name)  # set back the one that just clocked in so they dont have to find themself again
                on_select()

        ctime_var = tk.StringVar()
        ctime_label = tk.Label(ewin, textvariable=ctime_var, font=('Helvetica', 20), bg="white")
        ctime_label.place(relx=0.68, rely=0)

        def update_time():
            curtime = datetime.now().strftime("%H:%M:%S")
            ctime_var.set(curtime)
            ewin.update()
            ewin.after(500, update_time)  # update time every 0.5 seconds

        def ask_final_details():
            df = ('Helvetica', 14)
            question_pop = tk.Tk()
            question_pop.title('A few questions')
            question_pop.geometry("+700+400")
            info = dict()
            logger.debug("Asking for final details before sending report")

            def quit_pop_check():
                nonlocal info

                info = {"milage": milage.get(), "fuel_level": fuel_level.get(),
                        # "tanks_active_l": tanks_active_l.get(), "tanks_active_r": tanks_active_r.get(),
                        # "tanks_full_l": tanks_full_l.get(), "tanks_full_r": tanks_full_r.get(),
                        # "tanks_empty_l": tanks_empty_l.get(), "tanks_empty_r": tanks_empty_r.get(),
                        # "tank_pressure_l": tank_pressure_l.get(), "tank_pressure_r": tank_pressure_r.get()}
                        "tank_pressure": tank_pressure.get()}

                logger.debug(info.values())
                if all(info.values()):  # ask for final details while not all the info is there
                    question_pop.quit()
                    question_pop.destroy()

            def quit_pop():
                nonlocal info
                info = None
                logger.debug("No final details entered")
                question_pop.quit()
                question_pop.destroy()

            tk.Label(question_pop, text="Vehicle milage", font=df).grid(row=0, column=0, sticky=tk.E, padx=2)
            tk.Label(question_pop, text="Vehicle fuel level", font=df).grid(row=1, column=0, sticky=tk.E, padx=2)
            # Label(question_pop, text="#Tanks active L/R", font=df).grid(row=2, column=0, sticky=E, padx=2)
            # Label(question_pop, text="#Tanks full L/R", font=df).grid(row=3, column=0, sticky=E, padx=2)
            # Label(question_pop, text="#Tanks empty L/R", font=df).grid(row=4, column=0, sticky=E, padx=2)
            tk.Label(question_pop, text="Tank pressure", font=df).grid(row=2, column=0, sticky=tk.E,
                                                                       padx=2)  # used to be L/R on row 5
            milage = fts_widgets.ValidationEntry(question_pop, r"^[0-9]{0,7}$", width=10, font=df)
            milage.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=2)
            fuel_level = ttk.Combobox(question_pop,
                                      values=["Full", "7/8", "3/4", "5/8", "1/2", "3/8", "1/4", "1/8", "Empty"],
                                      state="readonly", width=9, font=df)
            fuel_level.grid(row=1, column=1, columnspan=2, sticky=tk.W)
            # TODO: remove tank pressure and just get it from the pi
            tank_pressure = fts_widgets.ValidationEntry(question_pop, r"^[0-9]{0,4}$", width=5, font=df)
            tank_pressure.grid(row=2, column=1, sticky=tk.W, pady=2, padx=2)
            tk.Button(question_pop, text="Continue", command=quit_pop_check, font=df).grid(row=3, column=0,
                                                                                           columnspan=3, sticky=tk.EW,
                                                                                           pady=5,
                                                                                           padx=2)  # used to be row 6
            # tanks_active_l = Entry(question_pop, width=5, font=df)
            # tanks_active_l.grid(row=2, column=1, sticky=W, pady=2, padx=2)
            # tanks_full_l = Entry(question_pop, width=5, font=df)
            # tanks_full_l.grid(row=3, column=1, sticky=W, pady=2, padx=2)
            # tanks_empty_l = Entry(question_pop, width=5, font=df)
            # tanks_empty_l.grid(row=4, column=1, sticky=W, pady=2, padx=2)
            # tank_pressure_l = Entry(question_pop, width=5, font=df)
            # tank_pressure_l.grid(row=5, column=1, sticky=W, pady=2, padx=2)
            # tanks_active_r = Entry(question_pop, width=5, font=df)
            # tanks_active_r.grid(row=2, column=2, sticky=W, pady=2, padx=2)
            # tanks_full_r = Entry(question_pop, width=5, font=df)
            # tanks_full_r.grid(row=3, column=2, sticky=W, pady=2, padx=2)
            # tanks_empty_r = Entry(question_pop, width=5, font=df)
            # tanks_empty_r.grid(row=4, column=2, sticky=W, pady=2, padx=2)
            # tank_pressure_r = Entry(question_pop, width=5, font=df)
            # tank_pressure_r.grid(row=5, column=2, sticky=W, pady=2, padx=2)

            question_pop.protocol("WM_DELETE_WINDOW", quit_pop)
            question_pop.mainloop()
            logger.debug(f"Final details entered: {info}")
            return info

        def send_report():
            nonlocal locked
            if locked:
                return
            locked = True
            if bulk_charge.has_ongoing():
                logger.debug("Still ongoing BULK")
                msg = "There is an ongoing BULK charge. You should deal with BULK charges before clocking out. Continue anyway?"
                if not messagebox.askyesno("Ongoing BULK", msg):
                    locked = False
                    return
            info = ask_final_details()
            if info is None:  # they hit the x
                locked = False
                return
            etime_var.set('Sending report...')
            employee_box.config(state=tk.DISABLED)
            clock_in_button.place_forget()
            send_report_button.place_forget()
            employee_name = employee_box.get()
            logger.debug("Sending final clock out request")
            request_params = {"type": "clock-done", "requester": str(Maint.mioskid),
                              "employee": employee_name}  # this option is only available if clocked in
            try:
                r = requests.get(url=constants.connect_url, params=request_params, timeout=2.0)
                inout = r.text
            except requests.exceptions.RequestException as e:
                logger.error("Clock out request failed", exc_info=e)
                return
            finally:
                employee_box.config(state='readonly')
                etime_var.set('Waiting for response...')
            if inout == "out":  # check current status
                logger.debug("Success, sending report email")

                clock_in_button.place(relx=0.13, rely=0.3)  # forget report, display clock in
                locked = False
                etime_var.set('Sending email...')  # otherwise send it
                ewin.update()
                while True:
                    try:
                        employee_email_sent = employee_log.send_report(Maint.mioskid, Maint.computername, info)
                    except Exception as e:
                        etime_var.set('Failed to send email. Unexpected error.')
                        logger.error("Employee email unexpected error", exc_info=e)
                        break
                    if employee_email_sent:
                        logger.info("Tech report sent")
                        etime_var.set("Sent report successfully")
                        guarantee_message_send("sent technician email")
                        self.last_clocked_in = None
                        ewin.destroy()
                        break
                    elif not tk.messagebox.askretrycancel("Email Send Error",
                                                          "The email could not be sent, but data will be saved. Try again?"):
                        etime_var.set('Failed to send email. Report saved.')
                        break
                os.remove(employee_log.LOG_FILE)  # so it is clear the technicican is completely clocked out
            else:
                # don't unlock, they'll have to close the window and try again
                etime_var.set('Server error.')  # didn't switch

        # def send_report():
        # 	Thread(target=t_send_report, daemon=True, name="sendreport").start()

        def t_clock_in():
            nonlocal locked
            if locked:
                return
            locked = True
            etime_var.set('Clocking in...')
            employee_box.config(state=tk.DISABLED)
            clock_in_button.place_forget()
            send_report_button.place_forget()
            employee_name = employee_box.get()
            request_params = {"type": "clock-toggle", "requester": str(Maint.mioskid),
                              "employee": employee_name}  # this option is only available if clocked out
            try:
                r = requests.get(url=constants.connect_url, params=request_params, timeout=2.0)
                inout = r.text
            except requests.exceptions.RequestException:
                return
            finally:
                locked = False
                etime_var.set('')
                employee_box.config(state='readonly')
            if inout == "in":  # check current status
                clock_out_button.place(relx=0.55, rely=0.3)  # forget clock in, display clock out
                if first_time:
                    employee_log.start_record(employee_name_map[employee_name])  # name -> id
                else:
                    employee_log.add_clock_in()
                ewin.destroy()  # on successful clock in, close the window because they always forget to close out of it and we end up with 50 windows open
                tk.messagebox.showinfo("Clock In", "Clocked in as " + employee_name)
            else:
                etime_var.set('Server error.')  # didn't switch

        def clock_in():
            threading.Thread(target=t_clock_in, daemon=True, name="clockin").start()

        def t_clock_out():
            nonlocal locked
            if locked:
                return
            locked = True
            etime_var.set('Clocking out...')
            employee_box.config(state=tk.DISABLED)
            clock_out_button.place_forget()
            employee_name = employee_box.get()
            request_params = {"type": "clock-toggle", "requester": str(Maint.mioskid),
                              "employee": employee_name}  # this option is only available if clocked in
            try:
                r = requests.get(url=constants.connect_url, params=request_params, timeout=2.0)
                inout = r.text
            except requests.exceptions.RequestException:
                return
            finally:
                locked = False
                etime_var.set('')
                employee_box.config(state='readonly')
            if inout == "break":  # check current status
                employee_log.add_clock_out()
                clock_out_button.place_forget()
                clock_in_button.place(relx=0.13, rely=0.3)  # forget clock out, display clock in
                send_report_button.place(relx=0.55, rely=0.3)  # allow to send report too
                # calculate delta-time
                clocked_time = int(employee_log.get_clockedin_time())  # total clocked-in time, in seconds
                mins, secs = divmod(clocked_time, 60)
                hours, mins = divmod(mins, 60)
                clocked_string = "Elapsed Time: {}:{:02}:{:02}".format(hours, mins, secs)
                etime_var.set(clocked_string)
            else:
                etime_var.set('Server error.')  # didn't switch

        def clock_out():
            threading.Thread(target=t_clock_out, daemon=True, name="clockout").start()

        clock_in_img = PhotoCache.get("clock_in.png")
        clock_out_img = PhotoCache.get("clock_out.png")
        send_report_img = PhotoCache.get("send_report.png")

        # start threads for connections so it doesn't lag the screen
        clock_in_button = tk.Button(ewin, image=clock_in_img, command=clock_in)
        clock_in_button.config(relief=tk.FLAT, highlightthickness=0, highlightcolor="white", bd=0,
                               activebackground="white")  # config to have invisible border at all times

        clock_out_button = tk.Button(ewin, image=clock_out_img, command=clock_out)
        clock_out_button.config(relief=tk.FLAT, highlightthickness=0, highlightcolor="white", bd=0,
                                activebackground="white")

        send_report_button = tk.Button(ewin, image=send_report_img, command=send_report)
        send_report_button.config(relief=tk.FLAT, highlightthickness=0, highlightcolor="white", bd=0,
                                  activebackground="white")

        etime_var = tk.StringVar()
        etime_label = tk.Label(ewin, textvariable=etime_var, font=('Helvetica', 20), bg="white")
        etime_label.place(relx=0.01, rely=0.80)

        update_time()

    def leak_test_pop(self):
        leak_win = tk.Toplevel(self.root)
        leak_win.attributes('-topmost', 'true')
        leak_win.title("Regulator and Leak Test")
        leak_win.geometry('400x300')

        leak_test_start_time = datetime.now()
        win_run = True
        phase = 0

        regular_font = ('Helvetica', 18)
        test_name = tk.StringVar()
        tk.Label(leak_win, textvariable=test_name, font=regular_font).place(relx=0.1, rely=0.05)
        current_pressure = tk.StringVar()
        tk.Label(leak_win, textvariable=current_pressure, font=('Helvetica', 24)).place(relx=0.5, rely=0.3,
                                                                                        anchor=tk.CENTER)
        time_left = tk.StringVar()
        tk.Label(leak_win, textvariable=time_left, font=('Consolas', 24)).place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        box_label = tk.Label(leak_win, text="Time:", font=("Helvetica", 16))
        box_label.place(relx=0.5, rely=0.7, anchor=tk.CENTER)
        test_time = tk.StringVar()
        test_time_box = tk.Spinbox(leak_win, textvariable=test_time, font=regular_font, from_=10, to=600, width=5,
                                   repeatinterval=20)
        test_time_box.place(relx=0.5, rely=0.8, anchor=tk.CENTER)
        test_time.set("60")

        def try_float(x):
            if x == "":
                return True
            try:
                float(x)
                return True
            except ValueError:
                return False

        n2_quality_box = fts_widgets.ValidationEntry(leak_win, validatecommand=try_float, font=('Consolas', 18),
                                                     width=5)

        def get_test_time():
            try:
                return int(test_time.get())
            except ValueError:
                return 60

        time_done = 0  # used for printing time left on the screen
        b_before = b_after = bh_before = bh_after = 0  # b = barrel, h = hose
        max_pressure = 0

        Pi.main.state_no_flow()
        Pi.safety.open()

        def update_pressure_loop():  # updates the main pressure display
            if win_run:
                current_pressure.set("{:.2f}".format(Pi.main.get_pressure_barrel().value))
                leak_win.update_idletasks()
                leak_win.update()
                leak_win.after(500, update_pressure_loop)

        update_pressure_loop()

        def update_time_loop():  # updates the time left display
            if win_run:
                time_amt_left = time_done - time.time()
                if time_amt_left > 0:
                    time_left.set(str(int(time_amt_left)))
                else:
                    time_left.set("")
                leak_win.after(200, update_time_loop)

        update_time_loop()

        # TODO: need to check HP safety that there is no leak
        #		probably as long as v1 is closed, check no decrease over long periods of time
        def start_leak_test_1():  # tests just the barrel
            nonlocal max_pressure, time_done, b_before
            max_pressure = Pi.main.get_pressure_barrelhose().value  # for regulator test
            Pi.main.state_no_flow()  # close hose off
            test_name.set("Barrel Leak Test")
            time.sleep(1.0)
            # record initial barrel pressure
            b_before = Pi.main.wait_for_steady_pressure(Pi.main.get_pressure_barrel, max_wait=10, max_deviation=0.15,
                                                        smoothing=1.0).value
            if not win_run:
                return
            audiohandler.play_wav("files/1.5-system-starting.wav")
            time_done = time.time() + get_test_time()  # set the timer
            leak_win.after(get_test_time() * 1000, end_leak_test_1)

        def end_leak_test_1():
            nonlocal b_after, phase
            b_after = Pi.main.get_pressure_barrel().value
            test_name.set("Let hose fill, then continue")
            audiohandler.play_wav("files/ding.wav")
            Pi.main.state_barrel_hose_inflate(block=False)
            phase = 2
            leak_win.after(4000,
                           setup_button_leak_test)  # wait a few seconds before letting them press to let barrel and hose inflate at least a little

        def start_leak_test_2():  # tests the barrel and hose
            nonlocal time_done, bh_before
            test_name.set("Barrel + Hose Leak Test")
            # hose will be open
            Pi.main.state_barrel_hose_flow()
            time.sleep(1)
            # record initial barrel+hose pressure
            bh_before = Pi.main.wait_for_steady_pressure(Pi.main.get_pressure_barrelhose, max_wait=15,
                                                         max_deviation=0.15, smoothing=1.0).value
            if not win_run:
                return
            audiohandler.play_wav("files/1.5-system-starting.wav")
            time_done = time.time() + get_test_time()  # set the timer
            leak_win.after(get_test_time() * 1000, end_leak_test_2)

        def end_leak_test_2():
            nonlocal bh_after, phase
            bh_after = Pi.main.get_pressure_barrelhose().value
            audiohandler.play_wav("files/ding.wav")
            phase = 3
            start_n2_test()  # kind of auto hit continue

        def start_n2_test():
            test_name.set("Enter Nitrogen Quality")
            box_label.config(text="N2 Quality %:")
            n2_quality_box.place(relx=0.5, rely=0.8, anchor=tk.CENTER)
            test_time_box.place_forget()
            main_button.config(text="Submit", state=tk.NORMAL)

        def end_n2_test():
            nonlocal phase
            n2_str = n2_quality_box.get()
            if n2_str != "":  # if empty, skip
                Maint.vals["nitrogen_percent"] = n2_str
                self.save_maint()
                # control message reporting nitrogen % (represented in Green on ctrl panel)
                msg = f"is currently serving ASTRAEA Nitrogen at {Maint.get_n2_percent():.1f}%"
                guarantee_message_send(msg)
            n2_quality_box.place_forget()
            box_label.place_forget()
            test_name.set("Finished All Tests")
            Pi.main.state_barrel_hose_deflate()  # release air
            phase = 4
            main_button.config(text="Print", state=tk.NORMAL)

        def setup_button_leak_test():
            main_button.config(text="Continue", state=tk.NORMAL)

        def start_regulator_test():
            nonlocal phase, leak_test_start_time
            test_name.set("Regulator Test")
            leak_test_start_time = datetime.now()
            Pi.main.state_barrel_inflate()
            phase = 1
            leak_win.after(1000, setup_button_leak_test)  # brief wait to avoid double-pressing

        def print_results():
            test_name.set("Printing...")

            byte_buffer = BytesIO()  # used as a write-capable file in memory
            pdf = PDFGen(byte_buffer, 2.75, 0.25)
            pdf.addtable(leak_test_start_time.strftime("%m/%d/%Y"), leak_test_start_time.strftime("%H:%M:%S"))
            pdf.addtable(Maint.computername, Maint.mioskid)
            pdf.options(15, True, align='center')
            pdf.addline("System Leak Test")
            pdf.skip(2)
            pdf.options(12)
            pdf.addtable("Max Pressure Ach.", "{:.2f} PSI".format(max_pressure))
            pdf.skip(5)

            letter_grades = ['A', 'B', 'C', 'D', 'F']
            f_grade = 0.05  # percent change needed for worst grade

            # keys are dpm, lrb, grade - Delta PSI/min, Leak Rate, Grade
            B_RESULTS = {}
            BH_RESULTS = {}

            def print_leak_result(title, before, after, store):
                pdf.options(14, True)
                pdf.addline(title)
                pdf.options(12)
                pdf.addtable("Start Pressure", "{:.2f} PSI".format(before))
                pdf.addtable("End Pressure", "{:.2f} PSI".format(after))
                dpm = (before - after) * 60 / get_test_time()
                # store the calculated data in the cache so it is not recomputed
                store['dpm'] = dpm
                store['lrb'] = 0
                store['grade'] = letter_grades[-1]  # worst
                pdf.addtable("Leak Rate", "{:.2f} PSI/min".format(max(dpm, 0)))
                if before != 0:
                    lrb = dpm / before
                    store['lrb'] = lrb
                    pdf.addtable("Leak Rate %", "{:.2f}%/min".format(max(lrb * 100, 0)))
                    gradenum = min(max(int((len(letter_grades) - 1) / f_grade * lrb), 0), len(letter_grades) - 1)
                    pdf.addtableadv("Grade", "{} ({})".format(letter_grades[gradenum], gradenum + 1), True, True)
                    store['grade'] = letter_grades[gradenum]

            # pass cache into calculators to store computed data
            print_leak_result("Barrel Leak Test", b_before, b_after, B_RESULTS)
            pdf.skip(5)
            print_leak_result("Barrel + Hose Leak Test", bh_before, bh_after, BH_RESULTS)

            # additional control message showing leak rates + grade (Represented in Green on Ctrl panel)
            msg = f"Results of Barrel Leak Test: {max(B_RESULTS['lrb'] * 100, 0):.2f} %/min leak rate (grade {B_RESULTS['grade']}). " \
                  f"Results of Barrel+Hose Leak Test: {max(BH_RESULTS['lrb'] * 100, 0):.2f} %/min leak rate (grade {BH_RESULTS['grade']})"
            guarantee_message_send(msg)

            pdf.skip(6)
            pdf.options(12, True)
            pdf.addtable("N2 Quality", f"{Maint.get_n2_percent():.1f}%")

            pdf.finish()

            byte_buffer.seek(0)
            convertandprint(byte_buffer)
            test_name.set("")
            main_button.config(text="Print", state=tk.NORMAL)

        def main_button_pressed():
            nonlocal phase
            main_button.config(state=tk.DISABLED)
            test_time_box.config(state=tk.DISABLED)
            if phase == 0:  # should say "start" then start the regulator test
                start_regulator_test()
            elif phase == 1:  # should say "continue" then continue to first leak test
                start_leak_test_1()
            elif phase == 2:  # should say "continue" then continue to next leak test
                start_leak_test_2()
            elif phase == 3:  # should say "Submit" then finish the n2 test
                end_n2_test()
            elif phase == 4:  # should say "print" then print
                print_results()

        def close():
            nonlocal win_run
            win_run = False
            Pi.main.stop()
            Pi.main.state_no_flow()
            Pi.safety.close()
            leak_win.destroy()

        main_button = tk.Button(leak_win, text="Start", font=regular_font, command=main_button_pressed)
        main_button.place(relx=0.95, rely=0.5, anchor=tk.E)

        # when they hit "End" or X out, run the close routine to shutdown valves and such
        tk.Button(leak_win, text="End", font=regular_font, command=close).place(relx=0.95, rely=0.9, anchor=tk.E)
        leak_win.protocol("WM_DELETE_WINDOW", close)

    def calib_pop(self):
        CalibrationWindow(self.root)

    def hpcalib_pop(self):
        pop = tk.Toplevel(self.root)
        pop.attributes('-topmost', True)
        pop.geometry("500x180+400+200")
        pop.transient(self.root)
        pop.title("HP Calibration")
        pop.grab_set()
        font = ("Helvetica", 14)

        raw_hp = Pi.main.get_hp(adjusted=False).value
        tk.Label(pop, text=f"Raw HP sensor reading: {int(raw_hp)}", font=font).pack(pady=5)
        tk.Label(pop, text=f"Adjusted HP reading: {int(Pi.main.get_hp().value)}", font=font).pack(pady=5)
        tk.Label(pop, text="HP Gauge:", font=font).pack(side=tk.LEFT, padx=5)
        entered_hp_var = tk.StringVar()
        fts_widgets.ValidationEntry(pop, r"|\d+", textvariable=entered_hp_var, font=font, width=8).pack(padx=2,
                                                                                                        side=tk.LEFT)

        def on_submit():
            try:
                hp = int(entered_hp_var.get())
            except ValueError:
                return

            Maint.vals['hp_scale'] = hp / raw_hp
            self.save_maint()
            pop.destroy()

        submit_btn = tk.Button(pop, text=" Submit ", font=font, command=on_submit)
        submit_btn.pack(side=tk.LEFT, padx=3)

    def tire_response_pop(self):
        pop = tk.Toplevel(self.root)
        pop.attributes('-topmost', True)
        pop.transient(self.root)
        pop.title("Tire Response Test")
        pop.grab_set()
        pop.update_idletasks()

        count_var = tk.StringVar()
        count = 0

        tk.Label(pop, text="N/A", textvariable=Maint.disp_pressure_S, padx=5, pady=5, font=("Consolas", 18)).grid(row=0,
                                                                                                                  column=0)
        tk.Label(pop, text="0", textvariable=count_var, padx=5, pady=5, font=("Consolas", 18)).grid(row=0, column=1)

        tk.Label(pop, text="Max Pressure:").grid(row=1, column=0)
        max_pressure_var = tk.StringVar()
        max_pressure_box = tk.Spinbox(pop, from_=10, to=150, increment=1, textvariable=max_pressure_var)
        max_pressure_box.grid(row=1, column=1)

        running = False
        barrel_data = []
        tire_data = []
        thread = None  # type: Optional[threading.Thread]

        Pi.main.state_barrel_hose_flow()

        def update_count():
            nonlocal count
            count += 1
            count_var.set(str(count))

        def inf_thread(max_pressure: int):
            nonlocal count
            count = 0

            tire_data.clear()
            barrel_data.clear()
            barrel_data.append(0)
            Pi.main.state_no_flow()
            Pi.safety.open()

            while running:
                # equalize
                Pi.main.state_barrel_hose_flow()
                tire_pressure = Pi.main.wait_for_steady_pressure(Pi.main.get_pressure_barrelhose, max_deviation=0.20)
                tire_data.append(tire_pressure.value)

                run_main_thread(update_count)

                if not running:
                    break
                if tire_pressure.value >= max_pressure:
                    run_main_thread(end)  # auto end
                    break

                # fill barrel with known amount of gas
                Pi.main.state_barrel_inflate(4)
                barrel_pressure = Pi.main.wait_for_steady_pressure(Pi.main.get_pressure_barrel, max_wait=5,
                                                                   max_deviation=0.35)
                barrel_data.append(
                    barrel_pressure.value - tire_pressure.value)  # amount it is filled with is proportial to

            Pi.main.state_no_flow()
            Pi.safety.close()

        def begin():
            nonlocal running, thread
            if thread and thread.is_alive():
                return
            running = True
            logger.info("Starting controlled inflate thread")
            try:
                max_pressure = int(max_pressure_var.get())
            except ValueError:
                return
            max_pressure_box.config(state=tk.DISABLED)
            thread = threading.Thread(target=inf_thread, name="inf-thread", args=[max_pressure])
            thread.start()

        def end():
            nonlocal thread, running
            running = False
            if thread and thread.is_alive():
                logger.info("Ending controlled inflate thread")
                thread.join(15)
                thread = None
            max_pressure_box.config(state=tk.NORMAL)
            show_graph()

        def show_graph():
            nonlocal barrel_data

            if not tire_data:
                logger.warning("No data to show")
                return

            # barrel data could have an extra datapoint if we didn't get to read the tire at that point
            if len(tire_data) != len(barrel_data):
                barrel_data = barrel_data[:-1]

            barrel_sum = np.cumsum(barrel_data)
            for x, y in zip(barrel_sum, tire_data):
                print(f"{x:.2f},{y:.2f}")

            fig, ax = plt.subplots()
            ax.plot(barrel_sum, tire_data, 'r-', label="Pressure (PSI)")  # red line
            ax.legend()
            plt.show()

        def close():
            nonlocal running
            running = False
            Pi.main.state_no_flow()
            Pi.safety.close()
            pop.destroy()

        tk.Button(pop, text="Begin", command=begin, padx=5, pady=5).grid(row=2, column=0)
        tk.Button(pop, text="End", command=end, padx=5, pady=5).grid(row=2, column=1)
        pop.protocol("WM_DELETE_WINDOW", close)

    def graphing_pop(self):
        pop = tk.Toplevel(self.root)
        pop.attributes('-topmost', True)
        pop.transient(self.root)
        pop.title("Graphing Controller")
        pop.update_idletasks()

        time_disp = tk.StringVar()
        pres_disp = tk.StringVar()
        temp_disp = tk.StringVar()

        tk.Label(pop, text="Time: N/A", textvariable=time_disp, padx=5, pady=5).pack()
        tk.Label(pop, text="Pres: N/A", textvariable=pres_disp, padx=5, pady=5).pack()
        tk.Label(pop, text="Temp: N/A", textvariable=temp_disp, padx=5, pady=5).pack()

        running = False
        thread_inst = None  # type: Optional[threading.Thread]
        interval = 5

        time_data = []
        pres_data = []
        temp_data = []
        start_time = 0

        def show_data(time_pt, pres_pt, temp_pt):
            time_disp.set(f"Time: {time_pt:6.2f}")
            pres_disp.set(f"Pres: {pres_pt:6.2f}")
            temp_disp.set(f"Temp: {temp_pt:6.2f}")

        def graph_thread():
            while running:
                # not exactly time-accurate running at 1/interval rate, but it really doesn't matter
                t1 = time.time() - start_time
                pres_data.append(Pi.main.get_pressure_barrelhose(smoothing=1 / interval).value)
                temp_data.append(Pi.main.get_temp_barrelhose(smoothing=1 / interval).value)
                t2 = time.time() - start_time
                time_data.append((t1 + t2) / 2)

                fts_util.run_main_thread(show_data, time_data[-1], pres_data[-1], temp_data[-1])

                time.sleep(1 / interval)

        def begin():
            nonlocal running, thread_inst, start_time
            if thread_inst and thread_inst.is_alive():
                return
            start_time = time.time()
            running = True
            thread_inst = threading.Thread(target=graph_thread, name="Graph-thread")
            logger.info("Starting graph thread")
            thread_inst.start()

        def end():
            nonlocal running, thread_inst
            if thread_inst and thread_inst.is_alive():
                logger.info("Ending graph thread")
                running = False
                thread_inst.join()
                thread_inst = None
            show_graph()

        def show_graph():
            if not time_data:
                logger.warning("No data to show")
                return

            mols_data = []
            for i in range(len(time_data)):
                pascals = pres_data[i] * 6894.76
                liters = 1  # arbitrary, but fixed
                kelvin = (temp_data[i] - 32) * 5 / 9 + 273.15
                moles = (pascals * liters) / (8.31446 * kelvin)  # gas constant included
                mols_data.append(moles * 0.1)  # scale

            fig, ax = plt.subplots()
            ax.plot(time_data, pres_data, 'r-', label="Pressure (PSI)")  # red line
            ax.plot(time_data, temp_data, 'b-', label="Temp (F)")  # blue line
            ax.plot(time_data, mols_data, 'g-', label="Moles x 10^-1")  # green line
            ax.legend()
            plt.show()

        tk.Button(pop, text="Begin", command=begin, padx=5, pady=5).pack()
        tk.Button(pop, text="End", command=end, padx=5, pady=5).pack()

    def tire_bp_cmd(self):
        tire_bp_cmd_window = tk.Toplevel(self.root)
        tire_bp_cmd_window.attributes('-topmost', True)
        tire_bp_cmd_window.transient(self.root)
        tire_bp_cmd_window.title("Tire BP [LOCKED]")
        tire_bp_cmd_window.update_idletasks()
        tire_bp_cmd_window.grab_set()

        locked = True
        tire_bp_cmd_window.geometry('300x180+1000+380')  # 300x180 screen at pos (1000, 380)

        instruct = tk.Label(tire_bp_cmd_window, text="Enter Code:")
        instruct.pack()
        enter_num = tk.Entry(tire_bp_cmd_window)
        enter_num.pack()

        def go():
            nonlocal locked
            if locked:
                keycode = enter_num.get()
                if keycode == "RKS" or keycode == "MAP":  # XXX: lmao
                    locked = False
                    tire_bp_cmd_window.title("Tire BP")
                    enter_num.delete(0, tk.END)
                    instruct.config(text="Enter Control Number:")
            else:
                try:
                    control_num = int(enter_num.get())
                except ValueError:
                    return
                logger.warning(f"TODO: print graph {control_num}")

        #
        # This is new
        #
        # try:
        # with open(JSON_JSON["min_statfile"]+"/log.file", 'a+') as f:
        # print("f.name",f.name)
        # for text in log_stream:
        # if text != '':
        # f.write(text + "|")
        # except Exception:
        # print("file log error", JSON_JSON["min_statfile"], "/log.file", " Invalid")
        # with open("log.file", 'a+') as f:
        # for text in log_stream:
        # if text != '':
        # f.write(text + "|")
        #
        # This is old
        #
        # with open("log.file") as f:
        # log = f.read()
        # lookfor = "^Control_Number. {:010}".format(control_num)
        # for line in log.split("\n"):
        # if re.match(lookfor, line):
        # print(line)
        # return

        submit_code = tk.Button(tire_bp_cmd_window, text="Submit", command=go)
        submit_code.pack()

        tire_bp_cmd_window.bind('<Return>', lambda e: submit_code.invoke())

    def fill_valve_timing(self):
        fill_valve_timing_window = tk.Toplevel()
        fill_valve_timing_window.attributes('-topmost', True)
        fill_valve_timing_window.update_idletasks()
        fill_valve_timing_window.title('fill_valve_timing_window')

        w = 640
        h = 480
        x = (self.SX / 2) - (w / 2)
        y = (self.SY / 2) - (h / 2)
        fill_valve_timing_window.geometry('%dx%d+%d+%d' % (w, h, x, y))
        title = tk.Label(fill_valve_timing_window, text='Valve cycle test', font=('Helvetica', 24))
        title.place(relx=0.3, rely=0.02)
        valve_run = False
        valve_cycle_count = 0

        Countvar = tk.StringVar()
        Countvar.set('count')

        runs_label = tk.Label(fill_valve_timing_window, font=("Helvetica", 20), textvariable=Countvar)
        runs_label.place(relx=0.5, rely=0.7)
        valve_thread = None  # type: Optional[threading.Thread]

        def valve_cycle():
            nonlocal valve_cycle_count
            valve_cycle_count = 0
            while valve_run:
                # takes 1.0 seconds to cycle through opening and closing all 5 valves
                for v in range(1, 6):
                    Pi.main.open(v)
                    time.sleep(0.1)
                    Pi.main.close(v)
                    time.sleep(0.1)
                valve_cycle_count += 1
                run_main_thread(Countvar.set, str(valve_cycle_count))

        def set_on():
            nonlocal valve_run, valve_thread
            if valve_thread is None or not valve_thread.is_alive():  # only start the thread if it is dead
                valve_run = True
                valve_thread = threading.Thread(target=valve_cycle, daemon=True, name="valvecycle")
                valve_thread.start()

        def set_off():
            nonlocal valve_run
            valve_run = False

        # valves will close when the next cycle is complete

        def kill_window():
            set_off()
            fill_valve_timing_window.destroy()

        V_run = tk.Button(fill_valve_timing_window, font=("Helvetica", 16), text='Run', command=set_on)
        V_run.place(relx=0.4, rely=0.5)

        V_stop = tk.Button(fill_valve_timing_window, font=("Helvetica", 16), text='Stop', command=set_off)
        V_stop.place(relx=0.6, rely=0.5)

        V_kill = tk.Button(fill_valve_timing_window, font=("Helvetica", 16), text='EXIT', command=kill_window)
        V_kill.place(relx=0.5, rely=0.8)

        fill_valve_timing_window.protocol("WM_DELETE_WINDOW", kill_window)

    def valve_control(self):
        V_cont_window = tk.Toplevel(self.root)
        V_cont_window.attributes('-topmost', True)
        V_cont_window.update_idletasks()
        V_cont_window.title('V_cont_window')

        ###SSH setup
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

        w = 640
        h = 480
        x = (self.SX / 2) - (w / 2)
        y = (self.SY / 2) - (h / 2)
        V_cont_window.geometry('%dx%d+%d+%d' % (w, h, x, y))

        ADC_pressures = [tk.StringVar() for _ in range(6)]
        ADC_temps = [tk.StringVar() for _ in range(8)]
        test_running = True

        def stop_test():
            nonlocal test_running
            test_running = False
            Pi.main.state_no_flow()
            V_cont_window.destroy()

        def read_ADCs():
            while test_running:
                # read pressures
                pressures = Pi.main.get_raw_pressures(smoothing=0)
                for ix in range(len(pressures)):
                    run_main_thread(ADC_pressures[ix].set, f"{pressures[ix].value:6.2f}")

                # read temps
                temps = Pi.main.get_temps(smoothing=0)
                for ix in range(min(len(temps), 8)):
                    run_main_thread(ADC_temps[ix].set, f"{temps[ix].value:6.1f}")

                time.sleep(0.1)

        threading.Thread(target=read_ADCs, daemon=True, name="readADCs").start()

        def pi_shutdown():  # Sends shutdown command to both raspberry pi's
            if pyro_run:
                ssh.connect('192.168.1.64', username='pi', password='raspberry')

                stdin, stdout, stderr = ssh.exec_command("wall shutdown called from control GUI")
                logger.debug(stdout.readlines())

                stdin, stdout, stderr = ssh.exec_command("sudo shutdown now -P -h")
                logger.debug(stdout.readlines())

        def check_safety():
            if pyro_run:
                messagebox.showinfo("PI Response", Pi.safety.diagnose())

        fa = tk.Frame(V_cont_window)
        fa.pack(fill=tk.X)

        faa = tk.Frame(V_cont_window)
        faa.pack(fill=tk.X)

        fb = tk.Frame(V_cont_window, height=64)
        fb.pack(fill=tk.BOTH, expand=1)

        fc = tk.Frame(V_cont_window, height=64)
        fc.pack(fill=tk.BOTH, expand=1)

        fd = tk.Frame(V_cont_window, height=64)
        fd.pack(side=tk.BOTTOM, fill=tk.BOTH, ipady=12)

        fe = tk.Frame(V_cont_window, height=64)
        fe.pack(side=tk.BOTTOM, fill=tk.BOTH, ipady=12)

        for i in range(6):
            tk.Label(fa, font=("Helvetica", 15), text=f"ADC{i + 1}").grid(row=0,
                                                                          column=i)  # ADC for analog-digital converter
            tk.Label(fa, font=("Consolas", 15), textvariable=ADC_pressures[i]).grid(row=1, column=i)
            fa.grid_columnconfigure(i, weight=1)  # expand grid to fill width

        # temperature sensors
        for i in range(8):
            tk.Label(faa, font=("Helvetica", 13), text=f"SR{i + 1}").grid(row=0,
                                                                          column=i)  # SR for serial (SPI for serial periferal interface)
            tk.Label(faa, font=("Consolas", 13), textvariable=ADC_temps[i]).grid(row=1, column=i)
            faa.grid_columnconfigure(i, weight=1)  # expand grid to fill width

        valve_threads = dict()

        # each valve gets its own thread. but not opening and closing the same valve. this is so you can't open a valve
        # then close it immediately before it even gets opened
        def safe_thread_run(label, targetfunc, open_valve):
            nonlocal valve_threads
            if label not in valve_threads or not valve_threads[
                label].is_alive():  # if the thread has never been started or it isnt alive
                vthread = threading.Thread(target=targetfunc, daemon=True, args=(open_valve,),
                                           name="controlvalve")  # start the thread
                vthread.start()
                valve_threads[label] = vthread  # and set its value in the function scope

        def control_valve(v, open_):
            if not Pi.main:
                return
            if open_:
                Pi.main.open(v)
            else:
                Pi.main.close(v)

        def tcontrol_v1(open_valve):
            control_valve(1, open_valve)
            V1O.config(bg="green" if open_valve else "white")
            V1C.config(bg="white" if open_valve else "red")

        def tcontrol_v2(open_valve):
            control_valve(2, open_valve)
            V2O.config(bg="green" if open_valve else "white")
            V2C.config(bg="white" if open_valve else "red")

        def tcontrol_v3(open_valve):
            control_valve(3, open_valve)
            V3O.config(bg="green" if open_valve else "white")
            V3C.config(bg="white" if open_valve else "red")

        def tcontrol_v4(open_valve):
            control_valve(4, open_valve)
            S1O.config(bg="green" if open_valve else "white")
            S1C.config(bg="white" if open_valve else "red")

        def tcontrol_vs(open_valve):
            control_valve(5, open_valve)
            S2O.config(bg="green" if open_valve else "white")
            S2C.config(bg="white" if open_valve else "red")

        # ensure all valves are closed to begin with
        ###Control Buttons

        V1O = tk.Button(fb, font=("Helvetica", 12), text='Open Valve 1',
                        command=partial(safe_thread_run, 'v1', tcontrol_v1, True))
        V1O.pack(side='left', expand=1)

        V2O = tk.Button(fb, font=("Helvetica", 12), text='Open Valve 2',
                        command=partial(safe_thread_run, 'v2', tcontrol_v2, True))
        V2O.pack(side='left', expand=1)

        V3O = tk.Button(fb, font=("Helvetica", 12), text='Open Valve 3',
                        command=partial(safe_thread_run, 'v3', tcontrol_v3, True))
        V3O.pack(side='left', expand=1)

        S1O = tk.Button(fb, font=("Helvetica", 12), text='Open Valve 4',
                        command=partial(safe_thread_run, 's1', tcontrol_v4, True))
        S1O.pack(side='left', expand=1)

        S2O = tk.Button(fb, font=("Helvetica", 12), text='Open Safety',
                        command=partial(safe_thread_run, 's2', tcontrol_vs, True))
        S2O.pack(side='left', expand=1)

        V1C = tk.Button(fc, font=("Helvetica", 12), text='Close Valve 1',
                        command=partial(safe_thread_run, 'v1', tcontrol_v1, False))
        V1C.pack(side='left', expand=1)

        V2C = tk.Button(fc, font=("Helvetica", 12), text='Close Valve 2',
                        command=partial(safe_thread_run, 'v2', tcontrol_v2, False))
        V2C.pack(side='left', expand=1)

        V3C = tk.Button(fc, font=("Helvetica", 12), text='Close Valve 3',
                        command=partial(safe_thread_run, 'v3', tcontrol_v3, False))
        V3C.pack(side='left', expand=1)

        S1C = tk.Button(fc, font=("Helvetica", 12), text='Close Valve 4',
                        command=partial(safe_thread_run, 's1', tcontrol_v4, False))
        S1C.pack(side='left', expand=1)

        S2C = tk.Button(fc, font=("Helvetica", 12), text='Close Safety',
                        command=partial(safe_thread_run, 's2', tcontrol_vs, False))
        S2C.pack(side='left', expand=1)

        safe_thread_run('v1', tcontrol_v1, False)
        safe_thread_run('v2', tcontrol_v2, False)
        safe_thread_run('v3', tcontrol_v3, False)
        safe_thread_run('v4', tcontrol_v4, False)
        safe_thread_run('vs', tcontrol_vs, False)

        ###Lower buttons

        shutdown_button = tk.Button(fe, font=("Helvetica", 16), text="Shutdown Pi's", command=pi_shutdown)
        shutdown_button.pack(side='left', expand=1)

        Plabel = tk.Label(fe, font=("Helvetica", 16), textvariable=Maint.disp_pressure_S)
        Plabel.pack(side='left', expand=1)

        start_control = tk.Button(fe, font=("Helvetica", 16), text="Check Safety", command=check_safety)
        start_control.pack(side='left', expand=1)

        V_cont_window.protocol("WM_DELETE_WINDOW", stop_test)

    def flowrate_test(self):
        pop = tk.Toplevel()
        pop.attributes('-topmost', 'true')
        pop.title("Flow Rate Benchmark")
        pop.geometry('300x220+1020+280')  # 300x220 screen at pos (1220, 280)

        # defaults, looking for a specific % resolution, so these will be adjusted automatically
        defaultres = 20  # %
        defaultresdec = defaultres / 100  # % as decimal
        ressigma = 0.25  # % as decimal

        font = ("Helvetica", 14)
        res_in = fts_widgets.LabeledEntry(pop, "Resolution %:", font=font)
        res_in.setentry(fts_widgets.NumberEntry, value=defaultres)
        res_in.pack()

        tk.Label(pop, font=("Consolas", 12), textvariable=Maint.disp_pressure_S).pack(pady=10)

        upresults = tk.StringVar()
        downresults = tk.StringVar()
        tk.Label(pop, font=font, textvariable=upresults).pack()
        tk.Label(pop, font=font, textvariable=downresults).pack()
        tk.Label(pop, font=font, text="1 second of valve opening results in score % barrel pressure change")

        running = False
        my_thread: Optional[threading.Thread] = None
        data = []
        Pi.safety.open()
        updt = downdt = 0

        def loop():
            nonlocal updt, downdt
            goingup = True
            res = res_in.entry.getvalue(defaultres) / 100  # input % as decimal
            upreslocked = False
            downreslocked = False
            updt = 1.0 * res / defaultresdec
            downdt = 0.5 * res / defaultresdec
            last_pressure = None
            while running and main_window.run:
                pressure = Pi.main.wait_for_steady_pressure(Pi.main.get_pressure_barrelhose,
                                                            max_wait=Maint.get_max_wait_time(), max_deviation=0.2).value

                if not (running and main_window.run):
                    break

                if last_pressure is not None:
                    # update the dt if necessary or lock it if within allowed error
                    if goingup and not upreslocked:
                        diff = abs(pressure - last_pressure) / (constants.regulator - last_pressure)
                        reserr = abs(diff - res) / res
                        print(f"{last_pressure} -> {pressure} = {diff:%}")
                        print(f"Error from desired resolution is {reserr:%}")
                        if reserr > ressigma:
                            print(updt)
                            updt *= res / diff
                            print(updt)
                        else:
                            upreslocked = True
                    elif not goingup and not downreslocked:
                        diff = abs(pressure - last_pressure) / last_pressure
                        reserr = abs(diff - res) / res
                        print(f"{last_pressure} -> {pressure} = {diff:%}")
                        print(f"Error from desired resolution is {reserr:%}")
                        if reserr > ressigma:
                            print(downdt)
                            downdt *= res / diff
                            print(downdt)
                        else:
                            downreslocked = True

                    # log the data if the dt is locked
                    if goingup and upreslocked or not goingup and downreslocked:
                        data.append((last_pressure, pressure))
                last_pressure = pressure

                if pressure > 130:
                    goingup = False
                elif pressure < 10:
                    goingup = True

                if goingup:
                    Pi.main.state_barrel_hose_inflate(updt)
                else:
                    Pi.main.state_barrel_hose_deflate(downdt)
            Pi.main.state_no_flow()

        def start():
            nonlocal running, my_thread
            running = True
            my_thread = threading.Thread(target=loop, name="flowrate test loop")
            my_thread.start()
            res_in.entry.config(state=tk.DISABLED)

        def stop():
            nonlocal running
            running = False
            Pi.main.stop()
            if not my_thread:
                return
            my_thread.join()
            if not data:
                return
            npdata = np.array(data)
            logger.debug(f"Inflate and delate data: {data}")
            befores = npdata[:, 0]
            afters = npdata[:, 1]
            diffs = afters - befores
            avgs = (befores + afters) / 2
            upi = diffs > 0
            downi = diffs < 0

            upeq = np.polyfit(avgs[upi], diffs[upi], 2)
            logger.info(f"Found up equation to be \n{upeq}")
            actual_reg = np.roots(upeq)[0]
            logger.info(f"Found effective regulator pressure to be {actual_reg:.1f}")
            ups = diffs[upi] / (actual_reg - avgs[upi])
            downs = diffs[downi] / avgs[downi]
            if len(ups) > 0:
                upmean = np.mean(ups)
                avgup = upmean / updt
                sigup = np.std(ups) / updt
                logger.info(f"Avg up was {upmean:.1%} / {updt:.2}s")
                upresults.set(f"Inflate score: {avgup:.1%} [{sigup:.1%}]")
                plt.scatter(befores[upi], afters[upi], c='red')
            if len(downs) > 0:
                downmean = np.mean(downs)
                avgdown = downmean / downdt
                sigdown = np.std(downs) / downdt
                logger.info(f"Avg down was {downmean:.1%} / {downdt:.2}s")
                downresults.set(f"Deflate score: {-avgdown:.1%} [{sigdown:.1%}]")
                plt.scatter(befores[downi], afters[downi], c='blue')
            plt.plot([0, actual_reg], [0, actual_reg], c='black')
            plt.show(block=False)

        startstop = fts_widgets.StateButton(pop, [("Start", start), ("Stop", stop)], font=font)
        startstop.pack()

        def close():
            if running:
                stop()
            Pi.safety.close()
            pop.destroy()

        pop.protocol("WM_DELETE_WINDOW", close)

    def change_coupon(self):
        coup_pop = tk.Toplevel()
        coup_pop.attributes('-topmost', 'true')
        coup_pop.title("Change Receipt Marketing Graphic")
        coup_pop.geometry('300x220+1220+280')  # 300x220 screen at pos (1220, 280)
        coupsamp = tk.Label(coup_pop, anchor=tk.S)

        Maint.valid_coupon_codes = {}
        current_code = tk.StringVar()
        current_code.set("None")
        rmg_parent = "\\\\APOLLO\\N24TyresMinionData\\{0}Miosk{1}\\{0}Miosk{1}RecptAds".format(Maint.computername,
                                                                                               Maint.mioskid)
        for rmgfile in os.listdir(rmg_parent):
            if rmgfile.endswith(".png") or rmgfile.endswith(".jpg"):
                Maint.valid_coupon_codes[os.path.splitext(rmgfile)[0]] = os.path.join(rmg_parent, rmgfile)

        def setphoto(x):
            pilp = Image.open(x)
            pilp.thumbnail((300, 180))  # scale to 0.4
            p = ImageTk.PhotoImage(image=pilp)  # must use PIL to open jpg
            coupsamp.config(image=p)
            coupsamp.image = p

        def chooseRMG():
            newrmg = filedialog.askopenfilename(initialdir=rmg_parent, title="Choose RMG",
                                                filetypes=(("Images", "png jpg"),))
            current_code.set(os.path.basename(newrmg).split(".")[0])
            go()

        if Maint.coupon_img != "":
            if Maint.coupon_img == "RANDOM":
                setphoto("files\\000RMG.jpg")
            else:
                setphoto(Maint.coupon_img)

        coupsamp.place(x=0, y=0, relwidth=1, relheight=1)  # set background
        tk.Entry(coup_pop, textvariable=current_code, width=9).place(relx=0.25, y=5)
        tk.Button(coup_pop, text="Select...", command=chooseRMG).place(relx=0.5, y=2)

        def go():
            testcode = current_code.get()
            if testcode in Maint.valid_coupon_codes.keys():
                Maint.coupon_img = Maint.valid_coupon_codes[testcode]
                setphoto(Maint.coupon_img)
                Maint.vals["rmg_code"] = testcode
            elif testcode == "000RMG":
                Maint.coupon_img = "RANDOM"
                setphoto("files\\000RMG.jpg")
                Maint.vals["rmg_code"] = "Random"
            else:
                Maint.coupon_img = ""
                coupsamp.config(image="")
                coupsamp.image = None
                Maint.vals["rmg_code"] = "(None)"

        submit_code = tk.Button(coup_pop, text="Save", command=go)
        submit_code.place(relx=0.7, y=2)

        coup_pop.bind('<Return>', lambda e: submit_code.invoke())

    def change_promo(self):
        change_promo_window = tk.Toplevel()
        change_promo_window.attributes('-topmost', 'true')
        locked = True
        change_promo_window.title("Promotional [LOCKED]")
        change_promo_window.geometry('300x180+1000+380')  # 300x180 screen at pos (1000, 380)

        instruct = tk.Label(change_promo_window, text="Enter Code:")
        instruct.pack()
        current_vars = tk.Label(change_promo_window)
        enter_title = tk.Entry(change_promo_window)
        enter_title.pack()
        enter_price = tk.Entry(change_promo_window)
        enter_price.config(width=6)

        def go():
            nonlocal locked
            if locked:
                keycode = enter_title.get()
                if keycode == "RKS" or keycode == "MAP":  # XXX: LMAO
                    locked = False
                    change_promo_window.title("Promotional")
                    enter_title.delete(0, tk.END)
                    instruct.config(text="Enter Title then Price")
                    submit_code.pack_forget()
                    enter_price.pack()
                    submit_code.pack()
                    promo_var_text = f'Current: "{Maint.promo_title}" at ${Maint.promo_price:.2f}'
                    current_vars.config(text=promo_var_text)
                    current_vars.pack()
            else:
                promo_title = enter_title.get()
                promo_price_local = enter_price.get()
                try:
                    promo_price = float(promo_price_local)
                except ValueError:
                    return
                logger.debug(f"set promo title/price {promo_title} @ ${promo_price:.2f}")
                submit_code.config(bg="green")
                promo_var_text = 'Current: "{}" at ${:.2f}'.format(promo_title, promo_price)
                current_vars.config(text=promo_var_text)

        submit_code = tk.Button(change_promo_window, text="Submit", command=go)
        submit_code.pack()

        change_promo_window.bind('<Return>', lambda e: submit_code.invoke())  # lambda takes an event e

    def gen_graph_popup(self):
        graph_pop = tk.Toplevel()
        graph_pop.attributes('-topmost', 'true')
        graph_pop.title("Lookup Inflation Graphs")
        graph_pop.geometry('300x450+1220+280')  # 300x400 screen at pos (1220, 280)
        instruct = tk.Label(graph_pop, text="Enter Control Number:", font=('Helvetica', 12))
        instruct.pack()
        ctrlnumenter = tk.StringVar()
        enter_ctrlnum = tk.Entry(graph_pop, font=('Helvetica', 20), textvariable=ctrlnumenter)
        enter_ctrlnum.config(width=11)
        enter_ctrlnum.pack()
        processtext = tk.Label(graph_pop, font=('Helvetica', 12))
        tire_buttons = []
        tire_names = []

        def check_ctrl():
            try:
                to_check = int(ctrlnumenter.get())
            except ValueError:
                enter_ctrlnum.config(bg="red")
                processtext.config(text="Invalid control number")
                return
            searchfor = "{:010}".format(to_check)
            filefound = None
            # check all minions available
            for checkdir in os.listdir(
                    r"\\APOLLO\N24TyresData\N24TyresMinionData"):  # contains each minion folder in format ?minionMiosk###
                if re.match(".*Miosk[0-9]{3}", checkdir):  # we can go inside and check
                    pathpath_search = os.path.join(r"\\APOLLO\N24TyresData\N24TyresMinionData", checkdir,
                                                   checkdir + "RAWData", "VehicleReceipts")
                    if os.path.exists(pathpath_search):  # some paths aren't complete yet
                        processtext.config(text="Searching " + checkdir)
                        graph_pop.update()
                        # this will list all files in the vehicle receipt folder.
                        for file in os.listdir(pathpath_search):
                            if file.startswith(searchfor) and file.endswith(".pdf"):
                                filefound = os.path.join(pathpath_search, file)
                                break
                if filefound:
                    break
            if filefound:
                try:
                    copyfile(filefound, 'CarReceipt.pdf')  # copy that file to local drive for printing
                except SameFileError:
                    pass  # just copied that file
                convertandprint('CarReceipt.pdf')
                enter_ctrlnum.config(bg="green")
                filename = os.path.basename(
                    filefound)  # just file, not path. should be ########## ddmmyy HHMMSS TIPS Vehicle Receipt.pdf
                justdate = filename[11:-25]  # shaves off everything unnessesary for date parsing
                try:
                    dateparse = datetime.strptime(justdate, "%d%m%y %H%M%S")
                    filedate = dateparse.strftime("%b %d, %I:%M %p")  # ex. Jan 12, 7:40 AM
                except Exception:
                    filedate = justdate
                processtext.config(text="File found from " + filedate)
            else:
                enter_ctrlnum.config(bg="orange")
                processtext.config(text="File not found")

        def get_graph(ctrlnum, tire, tsetpressure, filepath):  # setpressure can be None if old format
            logger.info("Getting PI graph for {:010}, {}".format(ctrlnum, tire))
            pigraph.generate_pi_graph(filepath, tsetpressure)

        def chk_graphs():
            try:
                to_check = int(ctrlnumenter.get())
            except ValueError:
                enter_ctrlnum.config(bg="red")
                processtext.config(text="Invalid control number")
                return
            searchfor = "{:010}".format(to_check)
            filefound = None
            # check all minions available
            for checkdir in os.listdir(
                    r"\\APOLLO\N24TyresData\N24TyresMinionData"):  # contains each minion folder in format ?minionMiosk###
                if re.match(".*Miosk[0-9]{3}", checkdir):  # we can go inside and check
                    pathpath_search = os.path.join(r"\\APOLLO\N24TyresData\N24TyresMinionData", checkdir,
                                                   checkdir + "RAWData",
                                                   "graphdata")  # "graphdata" must match what is in pigraph.py
                    if os.path.exists(pathpath_search):  # some paths aren't complete yet
                        processtext.config(text="Searching " + checkdir)
                        graph_pop.update()
                        # this will list all folders in the ata folder. There is probably a better way to search
                        for file in os.listdir(pathpath_search):
                            # make sure it's a folder (just in case)
                            if file.startswith(searchfor) and os.path.isdir(os.path.join(pathpath_search, file)):
                                filefound = os.path.join(pathpath_search, file)
                                break
                if filefound:
                    break
            tire_names.clear()  # tire names is list of tires that have graphs
            for b in tire_buttons:
                b.pack_forget()
            tire_buttons.clear()  # tire buttons is list of buttons that map to displaying that graph
            if filefound:
                processtext.config(text="Found")
                graph_pop.update()
                for tirefile in os.listdir(filefound):
                    tire_details = tirefile.split("-")  # file will look something like 0000002134-LFO-110.csv
                    if 2 <= len(tire_details) <= 3:  # possible formats
                        ctrlnum_compare = tire_details[0]
                        if ctrlnum_compare != searchfor:  # ensure correct control number
                            continue
                        tirename = tire_details[1]
                        if len(tire_details) == 3:  # new format includes setpressure
                            try:
                                tiresetpressure = int(tire_details[2][:-4])  # cuts off .csv
                            except ValueError as e:
                                logger.error(f"error parsing setpressure: {filefound} / {tirefile}", exc_info=e)
                                tiresetpressure = None
                        else:
                            tiresetpressure = None
                        logger.debug("got file. setpressure was " + str(tiresetpressure))
                        tire_names.append(tirename)
                        tire_buttons.append(tk.Button(graph_pop, text=tirename, font=('Helvetica', 15),
                                                      command=partial(get_graph, to_check, tirename, tiresetpressure,
                                                                      os.path.join(filefound, tirefile))))
                    else:
                        logger.error(f"error parsing setpressure: {filefound} / {tirefile}")

            if len(tire_buttons) > 0:
                enter_ctrlnum.config(bg="green")
                for b in tire_buttons:  # place all the buttons in a column
                    b.pack(fill=tk.BOTH)
            else:
                enter_ctrlnum.config(bg="white")
                processtext.config(text="No files found")

        submit_btn = tk.Button(graph_pop, text="See Graphs", font=('Helvetica', 14), command=chk_graphs)
        submit_btn.pack(pady=5)
        submit_btn = tk.Button(graph_pop, text="Print Receipt", font=('Helvetica', 14), command=check_ctrl)
        submit_btn.pack(pady=5)
        processtext.pack(anchor="sw")

    def gen_obc(self):
        obc_pop = tk.Toplevel()
        obc_pop.attributes('-topmost', 'true')
        obc_pop.title("Create Onboarding Code from Card")
        obc_pop.geometry('650x400+400+280')  # 650x400 screen at pos (400, 280)

        def force_caps(_e1, _e2, _e3):
            code_var.set(code_var.get().upper())

        code_var = tk.StringVar()
        tk.Label(obc_pop, font=('Helvetica', 18), text="Code:").grid(row=0, column=0, sticky=tk.E)
        tk.Entry(obc_pop, font=('Helvetica', 18), width=10, textvariable=code_var).grid(row=0, column=1, sticky=tk.W)
        code_var.trace_variable('w', force_caps)

        email_var = tk.StringVar()
        tk.Label(obc_pop, font=('Helvetica', 18), text="Email:").grid(row=1, column=0, sticky=tk.E)
        tk.Entry(obc_pop, font=('Helvetica', 18), width=20, textvariable=email_var).grid(row=1, column=1, sticky=tk.W)

        cof_var = tk.StringVar()
        tk.Label(obc_pop, font=('Helvetica', 18), text="COF:").grid(row=2, column=0, sticky=tk.E)
        tk.Entry(obc_pop, font=('Helvetica', 14), width=45, textvariable=cof_var, state="readonly").grid(row=2,
                                                                                                         column=1,
                                                                                                         sticky=tk.W)

        error_var = tk.StringVar()
        tk.Label(obc_pop, font=('Helvetica', 18), textvariable=error_var).grid(row=3, column=0, columnspan=2)

        oti = None  # type: Optional[otireader.OTIReader]

        def receive_preauth(auth: otireader.TransactionCompleteMessage):
            oti.send_cancel_transaction_msg()  # no matter, cancel the transaction, we just needed the ID
            if auth.status == Status.OK:
                if auth.authorization_details.transaction_db_id:
                    logger.debug("DB ID: " + auth.authorization_details.transaction_db_id)
                    cof = aprivatoken.get_cof(auth.authorization_details.transaction_db_id)
                    if cof:
                        logger.debug("COF: " + cof)
                        oti.send_show_message_msg("Success got COF", "************" + cof[-4:])
                        cof_var.set(cof)
                        error_var.set("Done")
                    else:
                        logger.warning("Failed to get COF")
                        cof_var.set("")
                        oti.send_show_message_msg("Failed to", "get COF")
                        error_var.set("Failure")
            # and then stop whether we got it or not
            elif auth.status == Status.CANCELLED:
                logger.info("Cancelled")
                oti.send_show_message_msg("Cancelled", "by user")  # and then stop
            elif auth.status == Status.TIMEOUT:
                logger.info("Timeout")
                oti.send_show_message_msg("Timeout")  # and then stop
            else:
                logger.info("Failed to read, Trying Again")
                oti.send_show_message_msg("Failed to read", "Try Again...")
                time.sleep(2)
                oti.send_pre_authorize_msg(COFPopup.PREAUTH_PRICE, callback=receive_preauth)
                return
            oti.stop()

        def swipe():
            nonlocal oti
            if oti and oti.is_listening():  # don't start another one
                return
            # 273D = heavy teardrop asterisk
            error_var.set("Please insert/swipe card or press the \u273D cancel.\n"
                          "Card will not be charged. This is for verification only.")
            oti = otireader.OTIReader(quiet=False)
            oti.send_pre_authorize_msg(COFPopup.PREAUTH_PRICE, callback=receive_preauth)

        def add_code():
            error_var.set("")
            email_add = email_var.get()
            if len(email_add) < 5 or '@' not in email_add:
                error_var.set("Invalid email")
                return
            code_add = code_var.get().upper()
            existing_contracts = Database.select("contracts", ["email"], code=code_add)
            if existing_contracts:
                msg = "Code already exists under email '" + existing_contracts[0]["email"] + \
                      "'. Are you sure you want to overwrite?"
                if not tk.messagebox.askokcancel("Warning", msg):
                    error_var.set("Not overwriting")
                    return
            cof_add = cof_var.get()
            if not cof_add.startswith("COF"):
                error_var.set("Invalid COF, please swipe card first.")
                return
            res = Database.insert_update("contracts", {'code': code_add, 'cofToken': cof_add, 'email': email_add},
                                         code=code_add)
            if res is None:
                error_var.set("Failed. Check connection.")
            else:
                error_var.set("Success.")

        def close():
            if oti and oti.is_listening():
                if tk.messagebox.askokcancel("OTI Warning", "We're still waiting for a response from the OTI reader.\n"
                                                            "Please click Cancel and respond to the reader. or\n"
                                                            "If you believe this is a mistake and want to exit, click OK.",
                                             icon='warning'):
                    logger.warning("OTI FORCE STOP")
                    oti.stop()
                else:
                    return  # cancel exit
            obc_pop.destroy()

        swipe_btn = tk.Button(obc_pop, text="Swipe Card", font=('Helvetica', 20), command=swipe)
        swipe_btn.grid(row=4, column=0, sticky=tk.NSEW)

        submit_btn = tk.Button(obc_pop, text="Create", font=('Helvetica', 20), command=add_code)
        submit_btn.grid(row=4, column=1, sticky=tk.NSEW)
        obc_pop.bind('<Return>', lambda e: submit_btn.invoke())  # lambda takes an event e
        obc_pop.protocol("WM_DELETE_WINDOW", close)

    def kiosk_info(self):
        t = tk.Toplevel()
        t.update_idletasks()  # Don't know why I should have to call this, but it's critical for the next line to work		t.overrideredirect(1)
        w = 250
        h = 100
        x = (self.SX / 2) - (w / 2)
        y = (self.SY / 2) - (h / 2)
        t.geometry('%dx%d+%d+%d' % (w, h, x, y))
        # t.geometry('128x64+{}+{}'.format(sx/2,sy/2))
        err = tk.Label(t, text="Name: " + str(Maint.computername))
        # err2 = Label(t, text="IP Address: " + str(public_ip))
        # convert uptime from seconds to days
        uptime_days = uptime() / 86400
        err3 = tk.Label(t, text="Time computer has been running: " + str(round(uptime_days, 1)) + " days")
        b = tk.Button(t, text='O.K.', command=t.destroy)
        err.pack()
        # err2.pack()
        err3.pack()
        b.pack(side=tk.BOTTOM)

        # tests
        e88 = tk.Entry(t)
        e88.pack()

    def bulk_charge_popup(self, root=None):
        pop = tk.Toplevel(root or self)
        pop.attributes('-topmost', 'true')
        pop.grab_set()  # technicians keeps forgetting to close the window, so this forces them to if they open it up randomly
        pop.title("Bulk Charge Manager")
        pop.geometry('600x150+920+250')  # 600x150 screen at pos (920, 250)
        status = tk.Label(pop, font=('Helvetica', 12))
        status.pack()

        def start_new():
            Maint.bulk_mode = True
            load()
            pop.after(500, pop.destroy)  # convenience for the technician

        def finish():
            cof, email, charges = bulk_charge.get_ongoing()
            total_amount = sum(x["amount"] for x in charges)
            if messagebox.askokcancel("Bill?", f"Ready to bill {email} for ${total_amount:.2f}?"):
                # FIXME: this is duplicated from gui.py
                logger.info(f"Charging {cof} for BULK {len(charges)} inflation(s): ${total_amount:.2f}")
                success, result = aprivatoken.charge_card(cof, total_amount, f"{len(charges)} inflation(s)")
                logger.info(result)
                if success and result and result["response_code"] == '0':  # APPROVAL:
                    guarantee_message_send(
                        f"Charged {email} (*{cof[-4:]}) for BULK {len(charges)} inflation(s): ${total_amount:.2f}")
                    finish_button.config(state=tk.DISABLED)  # if this happens to crash, don't let them charge again
                    trans = Data.Payment
                    # update cc params based on COF
                    trans.rdict = result  # success is already parsed by json in charge_card
                    trans.price_paid = Decimal(f"{trans.rdict['amount'] / 100.0:.2f}")
                    trans.transaction_id = trans.rdict["host_transaction_id"]
                    trans.pan = trans.rdict["tokenized_card_info"]["last_four"]
                    trans.status = trans.rdict["response_text"]
                    Data.COF.token = cof
                    Data.Contact.email = email
                    Data.Times.accept = datetime.now()
                    # tell control number fixing code to fix this uuid (get a good control number for it)
                    if Data.control_number is None:
                        Database.insert_into("__fix__", {"uid": Data.uuid, "controlNumber": None})
                    # "fix" all previously invalid transactions with the transaction id and PAN which paid for it
                    for charge in charges:
                        Database.insert_into("__bulk__", {"uid": charge["uid"], "transactionID": trans.transaction_id,
                                                          "pan": trans.pan})
                    Maint.allow_revenue_upload = True  # explicitly allow for bulk charges
                    Database.sync_all()  # now that the transactions have been fixed, sync the database to upload the bulk revenue data
                    receipt_bytes_pdf = genanyreceipt.create_bulk_receipt(charges)
                    threading.Thread(target=backup_files.save_bulk_receipt, name="save_bulk_receipt",
                                     args=(receipt_bytes_pdf.getvalue(), email, trans.price_paid)).start()
                    generated = convertpdf(receipt_bytes_pdf)
                    email_reciept.send_email([email], generated)
                    printimage(generated)
                    while messagebox.askyesno("Print receipt",
                                              "Print another?"):  # FIXME: this is like hitting a nail with a wrench
                        try:
                            printimage(generated)
                        except Exception:
                            continue
                    abort()  # abort now that charge is paid
                    pop.after(500, pop.destroy)  # convenience for the technician
                else:
                    # TODO: allow them to switch a COF to pay with
                    messagebox.showwarning("CC Fail",
                                           f"Card failed to charge.\nYou can call the office to have them charge instead, then clear the bulk.\n{result}")

        def abort():
            Maint.bulk_mode = False
            bulk_charge.clear()
            load()

        def ask_abort():
            if messagebox.askyesno("Abort?",
                                   "Are you sure you want to abort?\nThis will clear all recorded charges permenantly."):
                # remove unpaid (all) charges from revenue since they will never get sent
                cof, email, charges = bulk_charge.get_ongoing()
                total_amount = sum(x["amount"] for x in charges)
                logger.warning(f"BULK ABORT: {cof}, {email}, {total_amount:.2f} from {charges}")
                for charge in charges:
                    Database.delete("revenue", uid=charge["uid"])
                guarantee_message_send(
                    f"Aborting {email} (*{cof[-4:]}) BULK for {len(charges)} inflations: ${total_amount:.2f} [NOT CHARGED]",
                    'WARNING')
                abort()

        button_frame = tk.Frame(pop)
        button_frame.pack(pady=20)
        btn_font = ('Helvetica', 20)
        new_button = tk.Button(button_frame, text="Start New", font=btn_font, command=start_new)
        new_button.pack(side=tk.LEFT, padx=10)
        finish_button = tk.Button(button_frame, text="Finish (Bill)", font=btn_font, command=finish, state=tk.DISABLED)
        finish_button.pack(side=tk.LEFT, padx=10)
        abort_button = tk.Button(button_frame, text="Abort", font=btn_font, command=abort, state=tk.DISABLED)
        abort_button.pack(side=tk.LEFT, padx=10)

        def load():
            new_button.config(state=tk.DISABLED)
            finish_button.config(state=tk.DISABLED)
            abort_button.config(state=tk.DISABLED)
            if Maint.bulk_mode:
                abort_button.config(state=tk.NORMAL)  # 'abort' is only available in bulk mode
                if bulk_charge.has_ongoing():
                    cof, email, charges = bulk_charge.get_ongoing()
                    status.config(text=f"Bulk Charge ongoing for {email} (*{cof[-4:]}) #{len(charges) + 1}")
                    if len(charges) > 0:
                        finish_button.config(state=tk.NORMAL)  # cannot 'finish' if there are no inflations
                        abort_button.config(command=ask_abort)  # ask first if there are inflations
                else:
                    status.config(text="Bulk Charge mode is enabled")
            else:
                status.config(text=f"No ongoing Bulk Charge")
                new_button.config(state=tk.NORMAL)  # otherwise only 'new' available
            bulk_charge.display_charge_status_oti()

        load()
        return pop

    def restore_inflation_popup(self):
        pop = tk.Toplevel()
        pop.title("Restore Inflations Menu")
        pop.geometry('300x100+800+480')
        pop.transient(self.root)

        tk.Label(pop, text="You may be able to restore a prior inflation after a crash").pack()

        latest_bak_datestr = None
        pc_backups = os.listdir("pcbackups")
        for backup in pc_backups:
            # only check log.pc files
            if not backup.endswith("-log.pc"):
                continue

            # see if the date is the latest so far
            datestr = backup[:-7]
            if latest_bak_datestr is None or datestr > latest_bak_datestr:  # can be done using string comparison since date strings are the same length
                latest_bak_datestr = datestr

        # construct what the two files should look like
        latest_log_backup = f"pcbackups/{latest_bak_datestr}-log.pc"
        latest_pc_backup = f"pcbackups/{latest_bak_datestr}-latest.pc"

        def restore_last_inflation():
            shutil.copyfile(latest_log_backup, "log.pc")
            shutil.copyfile(latest_pc_backup, "latest.pc")
            messagebox.showinfo("Restart", "Program will now exit. Start the program again to restore the inflation.")
            main_window.close_window()

        if os.path.exists(latest_log_backup) and os.path.exists(latest_pc_backup):
            tk.Button(pop, text="Restore last inflation", command=restore_last_inflation).pack()
        else:
            tk.Label(pop, text="Sorry, not available").pack()
