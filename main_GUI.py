import os
from os import path

import tkinter as tk

from PIL import ImageTk, Image
import numpy as np
import subprocess
from functools import partial

from Tk_MsgBox.custom_msgbox import Ask_Msgbox, Info_Msgbox, Error_Msgbox, Warning_Msgbox
from Tk_Custom_Widget.ScrolledCanvas import ScrolledCanvas

from misc_module.image_resize import img_resize_dim, opencv_img_resize, pil_img_resize, open_pil_img
from misc_module.tk_img_module import *
from misc_module.tool_tip import CreateToolTip
from Report_Module.Report_GUI import Report_GUI
from WebSrc_Module.WebResource_GUI import WebResource_GUI

def widget_enable(*widgets):
    for widget in widgets:
        try:
            widget['state'] = 'normal'
        except AttributeError:
            pass

def widget_disable(*widgets):
    for widget in widgets:
        try:
            widget['state'] = 'disabled'
        except AttributeError:
            pass

class main_GUI():
    def __init__(self, master, window_icon = None):
        self.window_icon = window_icon

        self.master = master
        self.img_PATH = os.getcwd()
        self.load_TMS_logo()
        self.load_icon_src()

        self.top_frame = tk.Frame(master = self.master, bg = 'midnight blue'
            , highlightbackground = 'midnight blue')
        self.top_frame['height'] = 85
        self.top_frame.place(x = 0, y = 0
            , relwidth = 1, width = self.top_frame['width']
            , height = self.top_frame['height'], anchor = 'nw')
        self.top_frame.grid_rowconfigure(index = 0, weight = 1)
        self.top_frame.grid_rowconfigure(index = 1, weight = 1)
        self.top_frame.grid_columnconfigure(index = 3, weight = 1)

        tk_lb = tk.Label(self.top_frame, bg = 'white')
        tk_img_insert(tk_lb, self.tms_logo, img_width = 130
                            , img_height = 79
                            , pil_filter = Image.ANTIALIAS)

        tk_lb.grid(row = 0, column = 0, rowspan = 2, ipady = 1, ipadx = 1, sticky = 'nwse')

        self.menu_frame = tk.Frame(master = self.master, bg = 'blue'
            , highlightbackground = 'blue') #width = 44, height = 555
        self.menu_frame_w = 33 #44
        self.menu_frame['width'] = self.menu_frame_w
        self.menu_frame.place(x = 0, y = 85, relheight = 1, height = -85, anchor = 'nw')

        self.report_main_fr = ScrolledCanvas(master = self.master, frame_w = 970, frame_h = master.winfo_height() - 85
            , canvas_x = self.menu_frame_w, canvas_y = 85, bg = 'white'
            , hbar_x = self.menu_frame_w)

        self.resource_main_fr = ScrolledCanvas(master = self.master, frame_w = self.master.winfo_width()-self.menu_frame_w, frame_h = self.master.winfo_height()-85, 
            canvas_x = self.menu_frame_w, canvas_y = 85, bg = 'white'
            , hbar_x = self.menu_frame_w)

        self.__subframe_list = [self.report_main_fr
                                , self.resource_main_fr]

        main_GUI.class_report_gui = Report_GUI(self.report_main_fr.window_fr, scroll_canvas_class = self.report_main_fr
            , gui_graphic = dict( toggle_ON_btn_img = self.toggle_on_icon, toggle_OFF_btn_img = self.toggle_off_icon
                                , save_icon = self.save_icon, close_icon = self.close_icon
                                , up_arrow_icon = self.up_arrow_icon, down_arrow_icon = self.down_arrow_icon
                                , refresh_icon = self.refresh_icon
                                , text_icon = self.text_icon
                                , window_icon = self.window_icon
                                , folder_icon = self.folder_icon
                                , reset_icon = self.reset_icon
                                , clipboard_icon = self.report_b_icon 
                                )
            )
        main_GUI.class_report_gui.place(relwidth = 1, relheight =1, x=0,y=0)
        self.class_resource_gui = WebResource_GUI(self.resource_main_fr.window_fr, scroll_canvas_class = self.resource_main_fr)
        self.class_resource_gui.place(x = 0, y = 0, relx = 0, rely = 0, relwidth = 1, relheight = 1)

        self.report_ctrl_btn    = tk.Button(self.menu_frame, relief = tk.GROOVE, activebackground = 'navy', bg = 'royal blue')
        self.web_resource_btn   = tk.Button(self.menu_frame, relief = tk.GROOVE, activebackground = 'navy', bg = 'royal blue')

        tk_img_insert(self.report_ctrl_btn   , self.report_icon   , img_width = 26, img_height = 26)
        tk_img_insert(self.web_resource_btn  , self.web_icon      , img_width = 26, img_height = 26)


        self.__ctrl_btn_dict = {}
        hmap = self.__ctrl_btn_dict
        hmap[self.report_ctrl_btn]      = self.report_ctrl_btn_state
        hmap[self.web_resource_btn]     = self.resource_btn_state

        CreateToolTip(self.report_ctrl_btn, 'Report Generation'
            , 32, -5, font = 'Tahoma 11')
        CreateToolTip(self.web_resource_btn, 'Additional Resources'
            , 32, -5, font = 'Tahoma 11')

        self.report_ctrl_btn.grid(row = 0, column = 0, columnspan = 1, ipadx = 1, ipady = 1, sticky = 'nwse')
        self.web_resource_btn.grid(row = 1, column = 0, columnspan = 1, ipadx = 1, ipady = 1, sticky = 'nwse')

        self.report_ctrl_btn_state()

    def load_TMS_logo(self):
        self.tms_logo = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "Logo.png")

    def load_icon_src(self):
        self.report_icon     = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "clipboard (1).png")
        self.report_b_icon   = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "clipboard_black(1).png")
        self.web_icon        = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "web_resource.png")
        self.save_icon       = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "diskette.png")
        self.refresh_icon    = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "right.png")
        self.reset_icon      = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "reset.png")
        self.close_icon      = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "close.png")
        self.up_arrow_icon   = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "up_arrow.png")
        self.down_arrow_icon = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "down_arrow.png")
        self.text_icon       = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "text_icon.png")
        self.folder_icon     = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "folder.png")
        self.toggle_on_icon  = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "on icon.png")
        self.toggle_off_icon = open_pil_img(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "off icon.png")

    def __dummy_event(self, event):
        return "break"

    def ctrl_btn_state_func(self, target_btn):
        for ctrl_btn, btn_command in self.__ctrl_btn_dict.items():
            if ctrl_btn == target_btn:
                ctrl_btn['command'] = partial(self.__dummy_event, event = None)
                ctrl_btn['bg'] = 'navy'
            else:
                ctrl_btn['command'] = btn_command
                ctrl_btn['bg'] = 'royal blue'

    def show_subframe_func(self, target_frame, target_place = True, *args, **kwargs):
        for tk_frame in self.__subframe_list:
            if (isinstance(tk_frame, ScrolledCanvas)) == True:
                if tk_frame == target_frame:
                    if target_place == True:
                        tk_frame.show(*args, **kwargs)
                    elif target_place == False:
                        tk_frame.hide()

                else:
                    tk_frame.hide()

            elif (isinstance(tk_frame, tk.Frame)) == True:
                if tk_frame == target_frame:
                    if target_place == True:
                        tk_frame.place(*args, **kwargs)
                    elif target_place == False:
                        tk_frame.place_forget()

                else:
                    tk_frame.place_forget()

    def report_ctrl_btn_state(self):
        self.ctrl_btn_state_func(self.report_ctrl_btn)

        self.show_subframe_func(target_frame = self.report_main_fr, scroll_y = False)

        main_GUI.class_report_gui.btn_blink_start()

    def resource_btn_state(self):
        self.ctrl_btn_state_func(self.web_resource_btn)

        self.show_subframe_func(target_frame = self.resource_main_fr, scroll_x = False)

        main_GUI.class_report_gui.btn_blink_stop()
        

    def close_all(self):
        ask_msgbox = Ask_Msgbox('Do you want to quit?', title = 'Quit', parent = self.master, message_anchor = 'w')
        if ask_msgbox.ask_result() == True:
            self.master.destroy()
