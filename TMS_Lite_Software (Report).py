import os
from os import path

import sys

import tkinter as tk
from tkinter import ttk

from PIL import ImageTk, Image

from main_GUI import main_GUI

code_PATH = os.getcwd()
sys.path.append(code_PATH)

def create_save_folder(folder_dir = os.path.join(os.environ['USERPROFILE'],  "TMS_Temp_Save"), duplicate = False):
    if duplicate == True:
        if path.exists(folder_dir):
            index = 0
            loop = True
            while loop == True:
                new_dir = folder_dir + '({})'.format(index)
                if path.exists(new_dir):
                    index = index + 1
                else:
                    os.mkdir(new_dir)
                    loop = False

            return new_dir

        else:
            os.mkdir(folder_dir)
            return folder_dir
    else:
        if path.exists(folder_dir):
            #print ('File already exists')
            pass
        else:
            os.mkdir(folder_dir)
            #print ('File created')
        return folder_dir

class Root_Window(tk.Tk):
    def __init__(self,*args,**kw):
        tk.Tk.__init__(self,*args,**kw)
        self.withdraw() #hide the window
        self.after(0,self.deiconify) #as soon as possible (after app starts) show again

if __name__ == '__main__':
    main_window = Root_Window()
    main_window.title('TMS-Lite Software (Report) v.1.1.0')
    main_window.resizable(True, True)
    window_icon = ImageTk.PhotoImage(file = os.getcwd() + '\\TMS Icon\\' + 'logo4.ico')
    main_window_width = 890 #1080 #1280 #1080 #760       #890
    main_window_height = 600 #640 #720 #640 #720 #600    #600
    main_window.minsize(width=890, height=600)

    screen_width = main_window.winfo_screenwidth()
    screen_height = main_window.winfo_screenheight()

    x_coordinate = int((screen_width/2) - (main_window_width/2))
    y_coordinate = int((screen_height/2) - (main_window_height/2))

    main_window.geometry("{}x{}+{}+{}".format(main_window_width, main_window_height, x_coordinate, y_coordinate))
    main_window.iconphoto(False, window_icon)

    img_save_dir    = os.path.join(os.environ['USERPROFILE'],  "TMS_Saved_Images")
    report_save_dir = os.path.join(os.environ['USERPROFILE'], "TMS_Saved_Reports")
    create_save_folder(folder_dir = img_save_dir)
    create_save_folder(folder_dir = report_save_dir)

    main_panel = main_GUI(master = main_window, window_icon = window_icon)

    main_window.protocol("WM_DELETE_WINDOW", main_panel.close_all)
    main_window.mainloop()