import os
from os import path

import tkinter as tk
from tkinter import ttk

from Tk_MsgBox.custom_msgbox import Ask_Msgbox, Info_Msgbox, Error_Msgbox, Warning_Msgbox

from PIL import ImageTk, Image
import numpy as np
import subprocess
from functools import partial

from image_resize import img_resize_dim, opencv_img_resize, pil_img_resize, pil_icon_resize

from ScrolledCanvas import ScrolledCanvas

from Report_GUI import Report_GUI
from WebResource_GUI import WebResource_GUI

from tool_tip import CreateToolTip

def open_web_resource(file_name = "\\TMS - R&D Library.pdf"):
    err_flag = False
    folder_path = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__))) + "\\TMS_Web_Resources"
    
    file_path = folder_path + "\\" + file_name
    
    if path.isfile(file_path) == True:
        err_flag = False
        
        proc_obj = subprocess.Popen(['explorer', file_path]
            , stdout = subprocess.PIPE ## To show output values
            , stderr = subprocess.STDOUT ## To hide err values
            )

    elif path.isfile(file_path) == False:
        err_flag = True

    if err_flag == True:
        Error_Msgbox(message = 'Corresponding File Does Not Exist!', title = 'Error Web Resource(s)'
            , font = 'Helvetica 11', message_anchor = 'c')

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
        self.load_icon_img()

        self.top_frame = tk.Frame(master = self.master, bg = 'midnight blue'
            , highlightbackground = 'midnight blue', highlightthickness = 1) # width = 1080, height = 85
        self.top_frame['height'] = 85
        self.top_frame.place(relwidth = 1, relx = 0, rely =0)

        tk.Label(self.top_frame, image = self.tms_logo, bg = 'white').place(x=0, y=0)

        self.icon_frame = tk.Frame(master = self.master, bg = 'blue'
            , highlightbackground = 'blue', highlightthickness = 0) #width = 44, height = 555
        self.icon_frame_W = 33 #44
        self.icon_frame['width'] = self.icon_frame_W
        self.icon_frame.place(relheight = 1, x = 0, y = 85, height = -85)

        self.report_main_fr = ScrolledCanvas(master = self.master, frame_w = 970, frame_h = master.winfo_height() - 85, 
            canvas_x = self.icon_frame_W, canvas_y = 85, window_bg = 'white', canvas_highlightthickness = 0
            , hbar_x = self.icon_frame_W)
        self.report_main_fr.scrolly.place_forget()

        self.resource_main_fr = ScrolledCanvas(master = self.master, frame_w = self.master.winfo_width()-self.icon_frame_W, frame_h = self.master.winfo_height()-85, 
            canvas_x = self.icon_frame_W, canvas_y = 85, window_bg = 'white', canvas_highlightthickness = 0
            , hbar_x = self.icon_frame_W)

        self.__subframe_list = [self.report_main_fr
                                , self.resource_main_fr]

        main_GUI.class_report_gui = Report_GUI(self.report_main_fr.window_fr, scroll_canvas_class = self.report_main_fr
            , toggle_ON_btn_img = self.toggle_ON_button_img
            , toggle_OFF_btn_img = self.toggle_OFF_button_img
            , save_impil = self.save_impil
            , close_impil = self.close_impil
            , up_arrow_icon = self.up_arrow_icon
            , down_arrow_icon = self.down_arrow_icon
            , refresh_impil = self.refresh_impil
            , text_icon = self.text_icon
            , folder_impil = self.folder_impil
            , window_icon = self.window_icon)
        main_GUI.class_report_gui.place(relwidth = 1, relheight =1, x=0,y=0)


        self.class_resource_gui = WebResource_GUI(self.resource_main_fr.window_fr, scroll_canvas_class = self.resource_main_fr)
        self.class_resource_gui.place(x = 0, y = 0, relx = 0, rely = 0, relwidth = 1, relheight = 1)

        self.report_ctrl_btn = tk.Button(self.icon_frame, relief = tk.GROOVE, activebackground = 'navy', bg = 'royal blue', image=self.report_icon)

        self.web_resource_btn = tk.Button(self.icon_frame, relief = tk.GROOVE, activebackground = 'navy', bg = 'royal blue', image=self.web_resource_icon)
        
        self.__ctrl_btn_list = [self.report_ctrl_btn
                                , self.web_resource_btn]

        self.report_ctrl_btn['command'] = self.report_ctrl_btn_state
        self.web_resource_btn['command'] = self.resource_btn_state #open_web_resource

        CreateToolTip(self.report_ctrl_btn, 'Report Generation'
            , 32, -5, font = 'Tahoma 11')
        CreateToolTip(self.web_resource_btn, 'Additional Resources'
            , 32, -5, font = 'Tahoma 11')

        self.report_ctrl_btn.place(x= 0, y = 0)
        self.web_resource_btn.place(x= 0, y = 31)

        self.report_ctrl_btn_state()


    def load_TMS_logo(self):
        ### pil_filter, Image.: NEAREST, BILINEAR, BICUBIC and ANTIALIAS
        self.tms_logo, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "Logo.png"
            , img_width = 130
            , img_height = 79
            , pil_filter = Image.ANTIALIAS)

    def load_icon_img(self):
        self.report_icon, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "clipboard (1).png", img_width = 26, img_height =26)
        self.web_resource_icon, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "web_resource.png", img_width = 26, img_height =26)

        _, self.save_impil = pil_icon_resize(img_PATH = os.getcwd(), img_folder = "TMS Icon", img_file = "diskette.png", img_scale = 0.035)

        _, self.refresh_impil = pil_icon_resize(img_PATH = os.getcwd(), img_folder = "TMS Icon", img_file = "right.png", img_width = 18, img_height =18)

        _, self.close_impil = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "close.png", img_width = 20, img_height =20)

        self.up_arrow_icon, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "up_arrow.png", img_width = 20, img_height =20)
        self.down_arrow_icon, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "down_arrow.png", img_width = 20, img_height =20)

        self.text_icon, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "text_icon.png", img_width = 20, img_height =20)

        _, self.folder_impil = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "folder.png", img_width = 20, img_height =20)
        
        self.toggle_ON_button_img, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "on icon.png", img_scale = 0.06)
        self.toggle_OFF_button_img, _ = pil_icon_resize(img_PATH = self.img_PATH, img_folder = "TMS Icon", img_file = "off icon.png", img_scale = 0.06)

    def ctrl_btn_state_func(self, target_btn):
        for ctrl_btn in self.__ctrl_btn_list:
            if ctrl_btn == target_btn:
                widget_disable(ctrl_btn)
            else:
                widget_enable(ctrl_btn)

    def show_subframe_func(self, target_frame, target_place = True, *args, **kwargs):
        for tk_frame in self.__subframe_list:
            if (isinstance(tk_frame, ScrolledCanvas)) == True:
                if tk_frame == target_frame:
                    if target_place == True:
                        tk_frame.rmb_all_func(*args, **kwargs)
                    elif target_place == False:
                        tk_frame.forget_all_func()

                else:
                    tk_frame.forget_all_func()

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
            for widget in self.master.winfo_children(): # Loop through each widget in main window
                if isinstance(widget, tk.Toplevel): # If widget is an instance of toplevel
                    try:
                        widget.destroy()
                    except (tk.TclError):
                        pass
            self.master.destroy()
