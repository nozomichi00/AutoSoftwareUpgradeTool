import customtkinter
from tkinter import filedialog
import os
from time import sleep
from ctypes import windll
import shutil
import pyautogui
import cv2
import numpy as np
from datetime import datetime
from subprocess import run

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        # configure window
        self.title("Auto Software Upgrade Tool")
        self.iconbitmap('Images/AutoSoftwareUpgradeTool.ico')
        self.geometry(f"{800}x{600}")
        self.minsize(800, 600)

        # 設定3維框架排序
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        ##############################
        # Sidebar frame
        ##############################
        # Create sidebar frame
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")

        # Sidebar frame → Title
        self.sidebar_logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Menu bar", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.sidebar_logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Sidebar frame → Button
        self.sidebar_home_button = customtkinter.CTkButton(self.sidebar_frame, text="Home", command=self.home_button_event)
        self.sidebar_home_button.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_frame_2_button = customtkinter.CTkButton(self.sidebar_frame, text="Menu", command=self.frame_2_button_event)
        self.sidebar_frame_2_button.grid(row=2, column=0, padx=20, pady=10)
        self.sidebar_frame_3_button = customtkinter.CTkButton(self.sidebar_frame, text="Setting",command=self.frame_3_button_event)
        self.sidebar_frame_3_button.grid(row=3, column=0, padx=20, pady=10)

        # Sidebar frame → 分隔線
        self.sidebar_label = customtkinter.CTkLabel(self.sidebar_frame, text="")
        self.sidebar_label.grid(row=4, column=0, padx=20, pady=(20, 10))

        # Sidebar frame → Color Mode(Appearance)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Color Mode:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))

        # Sidebar frame → UI Scaling
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"], command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 10))

        # Sidebar frame → WT Scaling(Window Transparency)
        self.window_transparency_label = customtkinter.CTkLabel(self.sidebar_frame, text="WT Scaling:", anchor="w")
        self.window_transparency_label.grid(row=9, column=0, padx=20, pady=(10, 0))
        self.window_transparency_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["60%", "70%", "80%", "90%", "100%"], command=self.change_window_transparency_event)
        self.window_transparency_optionemenu.grid(row=10, column=0, padx=20, pady=(10, 10))
        
        ##############################
        # home frame
        ##############################
        # Create home frame
        self.home_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_rowconfigure(2, weight=1)
        self.home_frame.grid_columnconfigure(0, weight=1)
        
        # Status Motor
        self.step_status_motor_textbox = customtkinter.CTkTextbox(self.home_frame, height=30, activate_scrollbars=False, font=customtkinter.CTkFont(size=12, weight="bold"))
        self.step_status_motor_textbox.grid(row=0, column=0, padx=(20, 20), pady=(10, 0), sticky="nsew")
        self.detail_status_motor_textbox = customtkinter.CTkTextbox(self.home_frame, height=30, activate_scrollbars=False, font=customtkinter.CTkFont(size=12))
        self.detail_status_motor_textbox.grid(row=1, column=0, padx=(20, 20), pady=(10, 0), sticky="nsew")

        # create main frame
        self.main_frame = customtkinter.CTkFrame(self.home_frame)
        self.main_frame.grid(row=2, column=0, padx=(20, 20), pady=(10, 0), sticky="nsew")
        self.main_frame.grid_rowconfigure(26, weight=1)
        self.main_frame.grid_columnconfigure(4, weight=1)
        
        # Main frame → Step
        self.main_label_1 = customtkinter.CTkLabel(self.main_frame, text="Step 1", height=30)
        self.main_label_1.grid(row=1, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_2 = customtkinter.CTkLabel(self.main_frame, text="Step 2", height=30)
        self.main_label_2.grid(row=2, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_3 = customtkinter.CTkLabel(self.main_frame, text="Step 3", height=30)
        self.main_label_3.grid(row=3, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_4 = customtkinter.CTkLabel(self.main_frame, text="Step 4", height=30)
        self.main_label_4.grid(row=4, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_5 = customtkinter.CTkLabel(self.main_frame, text="Step 5", height=30)
        self.main_label_5.grid(row=5, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_6 = customtkinter.CTkLabel(self.main_frame, text="Step 6", height=30)
        self.main_label_6.grid(row=6, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_7 = customtkinter.CTkLabel(self.main_frame, text="Step 7", height=30)
        self.main_label_7.grid(row=7, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_8 = customtkinter.CTkLabel(self.main_frame, text="Step 8", height=30)
        self.main_label_8.grid(row=8, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_9 = customtkinter.CTkLabel(self.main_frame, text="Step 9", height=30)
        self.main_label_9.grid(row=9, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_label_10 = customtkinter.CTkLabel(self.main_frame, text="Step 10", height=30)
        self.main_label_10.grid(row=10, column=0, padx=(20, 0), pady=(10, 0), sticky="wn")
        
        # Main frame → File Path
        self.main_textbox_1 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_1.grid(row=1, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_2 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_2.grid(row=2, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_3 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_3.grid(row=3, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_4 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_4.grid(row=4, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_5 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_5.grid(row=5, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_6 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_6.grid(row=6, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_7 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_7.grid(row=7, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_8 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_8.grid(row=8, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_9 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_9.grid(row=9, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_textbox_10 = customtkinter.CTkTextbox(self.main_frame, width=300, height=30, activate_scrollbars=False)
        self.main_textbox_10.grid(row=10, column=1, padx=(20, 0), pady=(10, 0), sticky="wn")
        
        # Main frame → Select File
        self.main_button_1 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_1, self.main_textbox_1))
        self.main_button_1.grid(row=1, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_2 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_2, self.main_textbox_2))
        self.main_button_2.grid(row=2, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_3 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_3, self.main_textbox_3))
        self.main_button_3.grid(row=3, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_4 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_4, self.main_textbox_4))
        self.main_button_4.grid(row=4, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_5 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_5, self.main_textbox_5))
        self.main_button_5.grid(row=5, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_6 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_6, self.main_textbox_6))
        self.main_button_6.grid(row=6, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_7 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_7, self.main_textbox_7))
        self.main_button_7.grid(row=7, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_8 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_8, self.main_textbox_8))
        self.main_button_8.grid(row=8, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_9 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_9, self.main_textbox_9))
        self.main_button_9.grid(row=9, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")
        self.main_button_10 = customtkinter.CTkButton(self.main_frame, text="...", width=30, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.main_button_event(self.main_label_10, self.main_textbox_10))
        self.main_button_10.grid(row=10, column=2, padx=(20, 0), pady=(10, 0), sticky="wn")

        # Main frame → Backup Switch
        self.main_switch_1 = customtkinter.CTkSwitch(self.main_frame, text="Backup 1", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_1))
        self.main_switch_1.grid(row=1, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_2 = customtkinter.CTkSwitch(self.main_frame, text="Backup 2", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_2))
        self.main_switch_2.grid(row=2, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_3 = customtkinter.CTkSwitch(self.main_frame, text="Backup 3", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_3))
        self.main_switch_3.grid(row=3, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_4 = customtkinter.CTkSwitch(self.main_frame, text="Backup 4", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_4))
        self.main_switch_4.grid(row=4, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_5 = customtkinter.CTkSwitch(self.main_frame, text="Backup 5", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_5))
        self.main_switch_5.grid(row=5, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_6 = customtkinter.CTkSwitch(self.main_frame, text="Backup 6", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_6))
        self.main_switch_6.grid(row=6, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_7 = customtkinter.CTkSwitch(self.main_frame, text="Backup 7", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_7))
        self.main_switch_7.grid(row=7, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_8 = customtkinter.CTkSwitch(self.main_frame, text="Backup 8", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_8))
        self.main_switch_8.grid(row=8, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_9 = customtkinter.CTkSwitch(self.main_frame, text="Backup 9", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_9))
        self.main_switch_9.grid(row=9, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        self.main_switch_10 = customtkinter.CTkSwitch(self.main_frame, text="Backup 10", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.main_switch_10))
        self.main_switch_10.grid(row=10, column=3, padx=(20, 20), pady=(10, 0), sticky="wn")
        
        # Create operation frame
        self.operation_frame = customtkinter.CTkFrame(self.home_frame)
        self.operation_frame.grid(row=3, column=0, padx=(20, 20), pady=(10, 20), sticky="nsew")
        self.operation_frame.grid_rowconfigure(1, weight=1)
        self.operation_frame.grid_columnconfigure(4, weight=1)
        
        # operation frame → Start button
        self.operation_button_start = customtkinter.CTkButton(self.operation_frame, text="Start", width=70, height=30, font=customtkinter.CTkFont(weight="bold"), command=lambda: self.start_event(screen_switch, main_label, main_switch, main_textbox, self.operation_textbox))
        self.operation_button_start.grid(row=0, column=0, padx=(20, 20), pady=(10, 10), sticky="w")
        
        # operation frame → Threshold rate
        self.operation_label_1 = customtkinter.CTkLabel(self.operation_frame, text="Threshold rate(%):", height=30, font=customtkinter.CTkFont(weight="bold"))
        self.operation_label_1.grid(row=0, column=1, padx=(0, 0), pady=(10, 10), sticky="w")
        self.operation_textbox = customtkinter.CTkTextbox(self.operation_frame, width=35, height=30, activate_scrollbars=False, font=customtkinter.CTkFont(weight="bold"))
        self.operation_textbox.grid(row=0, column=1, padx=(130, 0), pady=(10, 10))
        
        # operation frame → Screen lock
        self.operation_label_2 = customtkinter.CTkLabel(self.operation_frame, text="Screen lock:", height=30, font=customtkinter.CTkFont(weight="bold"))
        self.operation_label_2.grid(row=0, column=2, padx=(13, 0), pady=(10, 10), sticky="w")
        self.operation_switch_win = customtkinter.CTkSwitch(self.operation_frame, text="Win", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.operation_switch_win))
        self.operation_switch_win.grid(row=0, column=3, padx=(7, 0), pady=(10, 10), sticky="w")
        self.operation_switch_tsmc = customtkinter.CTkSwitch(self.operation_frame, text="Tsmc", height=30, onvalue="on", offvalue="off", command=lambda: self.switch_event(self.operation_switch_tsmc))
        self.operation_switch_tsmc.grid(row=0, column=3, padx=(90, 0), pady=(10, 10))
        
        ##############################
        # second frame
        ##############################
        # Create second frame
        self.second_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.second_frame.grid_rowconfigure(1, weight=1)
        self.second_frame.grid_columnconfigure(0, weight=1)
        
        self.second_frame_label = customtkinter.CTkLabel(self.second_frame, text="Tool Manual", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.second_frame_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.second_frame_textbox = customtkinter.CTkTextbox(self.second_frame)
        self.second_frame_textbox.grid(row=1, column=0, padx=(20, 20), pady=(20, 20), sticky="nsew")
        self.second_frame_textbox.insert("0.0", '\
            1. 程式運行時，螢幕保護程式將被禁用，包含台積電的Screen Lock。\n\
            1. The screen saver will be disabled while the program is running.\n\
            \n\
            2. 程式運行時，“閾值率”被用來監控所有彈出視窗的匹配率(建議值70%~90%)。\n\
            2. "Threshold rate" is used to monitor the matching rate of all pop-ups during the procedure.\n\
                \tRecommanded value: 70%~90%\n\
            \n\
            3. 當 SystemBackup 被選中時，下面列出的操作將被執行：\n\
            3. Actions listed below will be execute when SystemBackup checked:\n\
            \t● backup\n\
            \t● Parameter\n\
            \t● xxx.log\n\
            \t● ver.txt\n\
            \t● Software backup\n\
            \n\
            4. “Step”文字的顏色:\n\
            4. The color of text "Step":\n\
            \t●Bold:\tFiles selected\n\
            \t●Orange:\tWorking...\n\
            \t●Green:\tSW Ver up finished\n\
            \t●Red:\tError occurred\n\
            \n\
            5. 程式運行時，請勿操作滑鼠游標和鍵盤，每個Step都將由Tool執行。\n\
            5. Do not operate mouse cursor, keyboard during Tool executing.\n\
            Every step will be executed by Tool.\n\
            \n\
            by Lucas 20221210')
        self.second_frame_textbox.configure(state="disabled")
        
        ##############################
        # third frame
        ##############################
        # Create third frame
        self.third_frame = customtkinter.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        self.button = customtkinter.CTkButton(self.third_frame, text="3")
        self.button.grid(row=1, column=0, padx=20, pady=10)

        ##############################
        # set default values
        ##############################
        # 預設框架顏色
        self.appearance_mode_optionemenu.set("System")
        # Modes: "System" (standard), "Dark", "Light"
        customtkinter.set_appearance_mode("System")
        # Themes: "blue" (standard), "green", "dark-blue"
        customtkinter.set_default_color_theme("blue")
        
        # 預設字體大小
        self.scaling_optionemenu.set("100%")
        
        # 預設透明度
        self.window_transparency_optionemenu.set("100%")
        
        # 預設 Switch
        self.main_switch_1.select()
        self.main_switch_1.configure(font=customtkinter.CTkFont(weight="bold"))
        self.operation_switch_win.select()
        self.operation_switch_win.configure(font=customtkinter.CTkFont(weight="bold"))
        self.operation_switch_tsmc.select()
        self.operation_switch_tsmc.configure(font=customtkinter.CTkFont(weight="bold"))
        
        # 預設 Monitor
        self.step_status_motor_textbox.insert("0.0", "Tool Running Status Monitor")
        self.detail_status_motor_textbox.insert("0.0", "Detail Monitor")
        self.operation_textbox.insert("0.0", "80")
        
        # 物件陣列
        screen_switch = [self.operation_switch_win, self.operation_switch_tsmc]
        main_label = [self.main_label_1, self.main_label_2, self.main_label_3, self.main_label_4, self.main_label_5,
                      self.main_label_6, self.main_label_7, self.main_label_8, self.main_label_8, self.main_label_10,]
        main_switch = [self.main_switch_1, self.main_switch_2, self.main_switch_3, self.main_switch_4, self.main_switch_5,
                       self.main_switch_6, self.main_switch_7, self.main_switch_8, self.main_switch_9, self.main_switch_10]
        main_textbox = [self.main_textbox_1, self.main_textbox_2, self.main_textbox_3, self.main_textbox_4, self.main_textbox_5,
                        self.main_textbox_6, self.main_textbox_7, self.main_textbox_8, self.main_textbox_9, self.main_textbox_10]
        
        # 執行 home frame
        self.select_frame_by_name("home")
        
    ##############################
    # Function
    ##############################
    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def change_window_transparency_event(self, new_window_transparency: str):
        new_window_transparency = float(new_window_transparency.replace("%", "")) / 100
        self.attributes("-alpha", new_window_transparency)

    def select_frame_by_name(self, name):
        # Change button weight
        if name == "home":
            self.sidebar_home_button.configure(font=customtkinter.CTkFont(weight="bold"))
        else:
            self.sidebar_home_button.configure(font=customtkinter.CTkFont(weight="normal"))
        if name == "frame_2":
            self.sidebar_frame_2_button.configure(font=customtkinter.CTkFont(weight="bold"))
        else:
            self.sidebar_frame_2_button.configure(font=customtkinter.CTkFont(weight="normal"))
        if name == "frame_3":
            self.sidebar_frame_3_button.configure(font=customtkinter.CTkFont(weight="bold"))
        else:
            self.sidebar_frame_3_button.configure(font=customtkinter.CTkFont(weight="normal"))

        # Show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.second_frame.grid_forget()
        if name == "frame_3":
            self.third_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.third_frame.grid_forget()
            
    def home_button_event(self):
        self.select_frame_by_name("home")

    def frame_2_button_event(self):
        self.select_frame_by_name("frame_2")

    def frame_3_button_event(self):
        self.select_frame_by_name("frame_3")
        
    def main_button_event(self, value1, value2):
        filename =  filedialog.askopenfilename(title = "Select file")
        value1.configure(font=customtkinter.CTkFont(weight="bold"))
        value2.insert("0.0", filename)
        
    def switch_event(self, value):
        if value.get() == "on":
            value.configure(font=customtkinter.CTkFont(weight="bold"))
        elif value.get() == "off":
            value.configure(font=customtkinter.CTkFont())
    
    def start_event(self, screen_switch, main_label, main_switch, main_textbox, threshold):
        ############################################################
        # Module1: Check before program execution.
        ############################################################
        # Get Windows Window Size
        screen_width, screen_height = pyautogui.size()
        
        # Windows Window Size Interlock
        if (screen_width != 1024 and screen_height != 768):
            
            # Get Tool Window Size
            win_widthm = self.winfo_width()
            win_height = self.winfo_height()
            
            # Window Move
            self.geometry("+" + str((screen_width - win_widthm) - 25) + "+" + str((screen_height - win_height) -80))
            self.update()
 
            # Set 7z.exe Path
            Un7zPath = os.path.abspath('./7z/7z.exe').replace('/', '\\\\')
            
            # Get Threshold rate
            threshold = (float(threshold.get("0.0", "end")) / 100)
            
            # AfterBackup Mode
            AfterBackup = True
            
            # Disable Windows screen saver
            if screen_switch[0].get == 'on':
                windll.kernel32.SetThreadExecutionState(0x80000002)
                self.detail_status_motor_textbox.delete("0.0", "end")
                self.detail_status_motor_textbox.insert("0.0", 'Windows screen saver has been successfully closed.')
                self.update()
                sleep(3)
            else:
                self.detail_status_motor_textbox.delete("0.0", "end")
                self.detail_status_motor_textbox.insert("0.0", 'The switch to disable the screen saver in Windows is not turned on.')
                self.update()
                sleep(3)
            
            # Disable Tsmc screen saver
            if screen_switch[1].get == 'on':
                try:
                    result = os.system("taskkill /f /im  VirtualScreenForm.exe")
                    if result == 0:
                        self.detail_status_motor_textbox.delete("0.0", "end")
                        self.detail_status_motor_textbox.insert("0.0", 'VirtualScreenForm.exe has been successfully closed.')
                        self.update()
                        sleep(3)
                    else:
                        self.detail_status_motor_textbox.delete("0.0", "end")
                        self.detail_status_motor_textbox.insert("0.0", 'VirtualScreenForm.exe is not running, so it does not need to be closed.')
                        self.update()
                        sleep(3)
                except:
                    self.detail_status_motor_textbox.delete("0.0", "end")
                    self.detail_status_motor_textbox.insert("0.0", 'An error occurred when closing VirtualScreenForm.exe, please check if this software is closed.')
                    self.update()
                    sleep(3)
            else:
                self.detail_status_motor_textbox.delete("0.0", "end")
                self.detail_status_motor_textbox.insert("0.0", 'The switch to disable the screen saver in Tsmc is not turned on.')
                self.update()
                sleep(3)
            
            # Main SW
            for i in range(0, 10):
                # Get UpgradeFiles Path
                UpgradeFileValue = main_textbox[i].get("0.0", "end")
                UpgradeFileValue = UpgradeFileValue[0:len(UpgradeFileValue) - 1].replace('/', '\\\\')

                # Get SystemBackup Mode Value
                BackupMode = main_switch[i].get()
                
                # Check UpgradeFiles Path
                if (UpgradeFileValue != ''):
                    # Tool Running Status Monitor
                    self.step_status_motor_textbox.delete("0.0", "end")
                    self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' UpgradeFile UnZip.')
                    main_label[i].configure(fg_color='#ffbf00')
                    self.update()
                    
                    # UpgradeFile UnZip
                    try:
                        # 檔案名稱、檔案路徑變數設定
                        UpgradeFileName = UpgradeFileValue[UpgradeFileValue.rfind('\\') + 1:len(UpgradeFileValue) - 4]
                        UpgradeFilePath = UpgradeFileValue[0:UpgradeFileValue.rfind('\\') + 1]
                        
                        # 第一層解壓縮
                        run([Un7zPath, "x", UpgradeFileValue, "-aoa", ("-o" + UpgradeFilePath + UpgradeFileName)])
                        
                        # 第二層解壓縮
                        run([Un7zPath, "x", (UpgradeFilePath + UpgradeFileName + '\\\\For_install\\\\InstallSoftware.EXE'), "-aoa", ("-o" + UpgradeFilePath + UpgradeFileName + '\\\\For_install')])

                        # Set UpgradeFilePath
                        UpgradeFile = (UpgradeFilePath + UpgradeFileName + '\\\\For_install\\\\Installer.exe')
                    except:
                        self.detail_status_motor_textbox.delete("0.0", "end")
                        self.detail_status_motor_textbox.insert("0.0", 'The UpgradeFile decompression failed, please confirm whether 7z.exe is missing.')
                        self.detail_status_motor_textbox.configure(text_color='red')
                        self.update()
                        break
                    
                    ############################################################
                    # Module2: Before System Backup
                    ############################################################
                    # Check BackupMode
                    if BackupMode == True:
                        self.step_status_motor_textbox.delete("0.0", "end")
                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' System Backup Mode.')
                        self.update()
                        
                        # Open UpgradeFile
                        if os.path.exists(UpgradeFile):
                            os.startfile(UpgradeFile)
                        else:
                            self.detail_status_motor_textbox.delete("0.0", "end")
                            self.detail_status_motor_textbox.insert("0.0", 'UpgradeFile not found, please check if decompression failed.')
                            self.detail_status_motor_textbox.configure(text_color='red')
                            self.update()
                            break
                        
                        # Copy Before SystemBackup Files
                        try:
                            # Set Day
                            currentDateAndTime = datetime.now()
                            if len(str(currentDateAndTime.day)) < 2:
                                currentDay = '0' + str(currentDateAndTime.day)
                            currentDay = str(currentDateAndTime.year) + str(currentDateAndTime.month) + currentDay
                            
                            # Check Folder
                            if not os.path.exists('C:\\BackUp'):
                                os.mkdir('C:\\BackUp')
                            BeforeDayFolder = 'C:\\BackUp\\' + str(currentDay) + "_Step"+ str(i + 1) +'_SystemBackup'
                            if not os.path.exists(BeforeDayFolder):
                                os.mkdir(BeforeDayFolder)
                            if os.path.exists(BeforeDayFolder + '\\EquipmentData\\Recipe'):
                                shutil.rmtree(BeforeDayFolder + '\\EquipmentData\\Recipe')
                            if os.path.exists(BeforeDayFolder + '\\EquipmentData\\Parameter'):
                                shutil.rmtree(BeforeDayFolder + '\\EquipmentData\\Parameter')
                                
                            # Copy Files
                            shutil.copy2('C:\\\\SuInstallHistory.log', BeforeDayFolder)
                            shutil.copy2('C:\\\\Bin\\ver.txt', BeforeDayFolder)
                            shutil.copytree('C:\\\\EquipmentData\\Recipe', (BeforeDayFolder + '\\EquipmentData\\Recipe'))
                            shutil.copytree('C:\\\\EquipmentData\\Parameter', (BeforeDayFolder + '\\EquipmentData\\Parameter'))
                        except:
                            self.detail_status_motor_textbox.delete("0.0", "end")
                            self.detail_status_motor_textbox.insert("0.0", 'Before System backup failed, please check if backup files are missing.')
                            self.detail_status_motor_textbox.configure(text_color='red')
                            self.update()
                            break
                        
                        # SystemBackup For Installer.exe
                        for j in range(1, 6):
                            # 設定變數
                            BackupStatus = True
                            Count = 0
                            
                            # 讀取檢查圖像，並轉換為 NumPy 陣列後再轉灰階
                            CheckImage = cv2.imread(('Images/CheckBackupImage_' + str(j) + '.jpg'), cv2.IMREAD_UNCHANGED)
                            CheckImage = cv2.cvtColor(CheckImage, cv2.COLOR_BGR2GRAY)
                            
                            # 匹配迴圈
                            while BackupStatus:
                                # 獲取當前桌面並轉換為 NumPy 陣列後再轉灰階
                                NowDesktop = pyautogui.screenshot()
                                NowDesktop = cv2.cvtColor(np.array(NowDesktop), cv2.COLOR_RGB2BGR)
                                NowDesktop = cv2.cvtColor(NowDesktop, cv2.COLOR_BGR2GRAY)
                                
                                # 計算匹配結果
                                CheckResult = cv2.matchTemplate(NowDesktop, CheckImage, cv2.TM_CCOEFF_NORMED)
                                
                                # 取得匹配成功的參數
                                _, max_val, _, _ = cv2.minMaxLoc(CheckResult)

                                # Check result threshold rate(%) 
                                yloc, xloc = np.where(CheckResult >= threshold)

                                # Detail Monitor
                                self.detail_status_motor_textbox.delete("0.0", "end")
                                self.detail_status_motor_textbox.insert("0.0", 'Check: ' + str(j) + '/5, Elapsed time: ' + str(Count) + '/s, Threshold rate: ' + str(max_val)[2:4] + '%/' + str(threshold * 100)[0:2] + '%')
                                self.update()
                                
                                # Check Array Count
                                if len(yloc) != 0:
                                    if j == 1:
                                        # 按備份
                                        pyautogui.click((xloc[0] + 200), (yloc[0] + 20), 3)
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' System Backup Button.')
                                        self.update()
                                        
                                    elif j == 2:
                                        # 按同意
                                        pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' System Backup Start.')
                                        self.update()
                                        
                                    elif j == 3:
                                        # 檢查是否有開始備份
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' Now is system backup.')
                                        self.update()
                                        
                                    elif j == 4:
                                        # 按完成
                                        pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' System Backup Done.')
                                        self.update()
                                        
                                    elif j == 5:
                                        # 按關閉重啟
                                        pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                        sleep(2)
                                        
                                        # Open UpgradeFile
                                        if os.path.exists(UpgradeFile):
                                            os.startfile(UpgradeFile)
                                        else:
                                            self.detail_status_motor_textbox.delete("0.0", "end")
                                            self.detail_status_motor_textbox.insert("0.0", 'UpgradeFile not found, please check if decompression failed.')
                                            self.detail_status_motor_textbox.configure(text_color='red')
                                            self.update()
                                            break
                                        
                                    # Close while loop.
                                    BackupStatus = False
                                else:
                                    sleep(1)
                                    Count = Count + 1
                    else:
                        # Open UpgradeFile
                        if os.path.exists(UpgradeFile):
                            os.startfile(UpgradeFile)
                        else:
                            self.detail_status_motor_textbox.delete("0.0", "end")
                            self.detail_status_motor_textbox.insert("0.0", 'UpgradeFile not found, please check if decompression failed.')
                            self.detail_status_motor_textbox.configure(text_color='red')
                            self.update()
                            break
                    
                    ############################################################
                    # Module3: Software Upgrade
                    ############################################################
                    # System Upgrade
                    for k in range(1, 7):
                        UpgradeStatus = True
                        Count = 0
                            
                        CheckImage = cv2.imread(('Images/CheckUpgradeImage_' + str(k) + '.jpg'), cv2.IMREAD_UNCHANGED)
                        CheckImage = cv2.cvtColor(CheckImage, cv2.COLOR_BGR2GRAY)
                        
                        while UpgradeStatus:
                            NowDesktop = pyautogui.screenshot()
                            NowDesktop = cv2.cvtColor(np.array(NowDesktop), cv2.COLOR_RGB2BGR)
                            NowDesktop = cv2.cvtColor(NowDesktop, cv2.COLOR_BGR2GRAY)
                            CheckResult = cv2.matchTemplate(NowDesktop, CheckImage, cv2.TM_CCOEFF_NORMED)
                            _, max_val, _, _ = cv2.minMaxLoc(CheckResult)
                            yloc, xloc = np.where(CheckResult >= threshold)

                            self.detail_status_motor_textbox.delete("0.0", "end")
                            self.detail_status_motor_textbox.insert("0.0", 'Check: ' + str(k) + '/6, Elapsed time: ' + str(Count) + '/s, Threshold rate: ' + str(max_val)[2:4] + '%/' + str(threshold * 100)[0:2] + '%')
                            self.update()

                            if len(yloc) != 0:
                                if k == 1:
                                    # 解除System Backup Mode
                                    pyautogui.click(xloc[0], yloc[0], 3)
                                    for l in range(1, 10):
                                        pyautogui.press('ctrl')
                                        sleep(0.1)
                                    self.step_status_motor_textbox.delete("0.0", "end")
                                    self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' SW Upgrade Unlock.')
                                    self.update()
                                        
                                elif k == 2:
                                    # 按升級
                                    pyautogui.click((xloc[0] + 200), (yloc[0] + 20), 3)
                                    self.step_status_motor_textbox.delete("0.0", "end")
                                    self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' SW Upgrade Button.')
                                    self.update()
                                    
                                elif k == 3:
                                    # 按同意
                                    pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                    self.step_status_motor_textbox.delete("0.0", "end")
                                    self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' SW Upgrade Start.')
                                    self.update()
                                    
                                elif k == 4:
                                    # 檢查是否有開始備份
                                    self.step_status_motor_textbox.delete("0.0", "end")
                                    self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' Now is system update.')
                                    self.update()
                                    
                                elif k == 5:
                                    # 按完成
                                    pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                    self.step_status_motor_textbox.delete("0.0", "end")
                                    self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' SW Upgrade Done.')
                                    self.update()
                                    
                                elif k == 6:
                                    # 按關閉
                                    pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                    self.update()
                                    sleep(2)
                                    
                                UpgradeStatus = False
                            else:
                                sleep(1)
                                Count = Count + 1
                        
                    # Change step label color
                    main_label[i].configure(fg_color='green')
                    self.update()   
                ############################################################
                # Module4: After System Backup
                ############################################################
                elif AfterBackup:
                    if ((i + 1) == 1):
                        # Tool Running Status Monitor
                        self.step_status_motor_textbox.delete("0.0", "end")
                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' After System Backup.')
                        self.detail_status_motor_textbox.delete("0.0", "end")
                        self.detail_status_motor_textbox.insert("0.0", 'Please select an upgrade file(ISU32-*_Installer_00.exe)')
                        self.detail_status_motor_textbox.configure(text_color='red')
                        main_label[i].configure(fg_color='red')
                        self.update()
                        break
                    else:
                        # Tool Running Status Monitor
                        self.step_status_motor_textbox.delete("0.0", "end")
                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i + 1) + ' After System Backup.')
                        self.update()
                        
                        # Open Backup.exe
                        if os.path.exists('c:\\\\BackUp\\\\Backup.exe'):
                            os.startfile('c:\\\\BackUp\\\\Backup.exe')
                        else:
                            self.detail_status_motor_textbox.delete("0.0", "end")
                            self.detail_status_motor_textbox.insert("0.0", 'Backup.exe not found, please perform After System Backup manually.')
                            self.detail_status_motor_textbox.configure(text_color='red')
                            self.update()
                            break
                        
                        # SystemBackup For Backup.exe
                        for m in range(1, 6):
                            BackupStatus = True
                            Count = 0
                            
                            CheckImage = cv2.imread(('Images/CheckBackupImage_' + str(m) + '.jpg'), cv2.IMREAD_UNCHANGED)
                            CheckImage = cv2.cvtColor(CheckImage, cv2.COLOR_BGR2GRAY)
                            
                            while BackupStatus:
                                NowDesktop = pyautogui.screenshot()
                                NowDesktop = cv2.cvtColor(np.array(NowDesktop), cv2.COLOR_RGB2BGR)
                                NowDesktop = cv2.cvtColor(NowDesktop, cv2.COLOR_BGR2GRAY)
                                CheckResult = cv2.matchTemplate(NowDesktop, CheckImage, cv2.TM_CCOEFF_NORMED)
                                _, max_val, _, _ = cv2.minMaxLoc(CheckResult)
                                yloc, xloc = np.where(CheckResult >= threshold)
                                
                                self.detail_status_motor_textbox.delete("0.0", "end")
                                self.detail_status_motor_textbox.insert("0.0", 'Check: ' + str(m) + '/5, Elapsed time: ' + str(Count) + '/s, Threshold rate: ' + str(max_val)[2:4] + '%/' + str(threshold * 100)[0:2] + '%')
                                self.update()
                                
                                if len(yloc) != 0:
                                    if m == 1:
                                        # 按備份
                                        pyautogui.click((xloc[0] + 200), (yloc[0] + 20), 3)
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i) + ' After SystemBackup Button.')
                                        self.update()
                                        
                                    elif m == 2:
                                        # 按同意
                                        pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i) + ' After SystemBackup Start.')
                                        self.update()

                                    elif m == 3:
                                        # 檢查是否有開始備份
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i) + ' Now is After SystemBackup.')
                                        self.update()
                                        
                                    elif m == 4:
                                        # 按完成
                                        pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i) + ' After SystemBackup Done.')
                                        self.update()
                                        
                                    elif m == 5:
                                        # 按關閉
                                        pyautogui.click((xloc[0] + 40), (yloc[0] + 15), 3)
                                        self.step_status_motor_textbox.delete("0.0", "end")
                                        self.step_status_motor_textbox.insert("0.0", 'Step:' + str(i) + ' Software upgrade completed.')
                                        self.update()

                                        AfterBackup = False
                                        break
                                        
                                    BackupStatus = False
                                else:
                                    sleep(1)
                                    Count = Count + 1
                # Next step no path.
                else:
                    break
        # 解析度驗證失敗
        else:
            self.step_status_motor_textbox.delete("0.0", "end")
            self.step_status_motor_textbox.insert("0.0", 'Window size Error')
            self.detail_status_motor_textbox.delete("0.0", "end")
            self.detail_status_motor_textbox.insert("0.0", 'Window size is not 1024 x 768.')
            self.detail_status_motor_textbox.configure(text_color='red')

if __name__ == "__main__":
    app = App()
    app.mainloop()

