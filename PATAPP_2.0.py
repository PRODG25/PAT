#Download libraries
import tkinter as tk
from tkinter import ttk
import subprocess
import pyautogui
import time
import win32com.client
import keyboard
import easygui
import pygetwindow as gw
import pyautogui
import win32com.client as win32
import win32gui
import win32con
import re
import webbrowser
import sys
import pandas as pd
import os
import ctypes
from PIL import Image, ImageTk
from datetime import datetime
import openpyxl

#Display Scaling and Resolution check
width, height = pyautogui.size()
if width != 1920 or height != 1080:
    # Display a message box
    ctypes.windll.user32.MessageBoxW(
        0,
        "Please change the display resolution to 1920x1080 on your desktop and restart the app.",
        "Resolution Error",
        0x30 | 0x0  # MB_ICONWARNING | MB_OK
    )
    sys.exit()

scaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
if scaleFactor != 1.5:
    # Display a message box
    ctypes.windll.user32.MessageBoxW(
        0,
        "Please change the display scaling to 150% (1.5x) on your desktop and restart the app.",
        "Scaling Error",
        0x30 | 0x0  # MB_ICONWARNING | MB_OK
    )
    sys.exit()

def button_click(button_number):
    if button_number == 1:
        custom_x10()
        step10()
        result_label.config(text="Flow for step 10 completed")
    elif button_number == 2:
        custom_x10()
        step15()
        result_label.config(text="Redirected to Cloud Flow for step 15")
    elif button_number == 3:
        custom_x10()
        step41()
        result_label.config(text="Redirected to Cloud Flow for step 41")
    elif button_number == 4:
        custom_x10()
        step7wwf()
        result_label.config(text="Redirected to Cloud Flow for 7 week old flows")
    elif button_number == 5:
        custom_x10()
        step8wwf()
        result_label.config(text="Flow for 8 week completed")
    elif button_number == 6:
        custom_x10()
        refresh_data()  # When "Big Button" is clicked, refresh data
        result_label.config(text="Data Refreshed!")
    elif button_number == 7:
        open_excel_file()  # When "Open Excel" button is clicked, open Excel file
        result_label.config(text="Excel File Opened")
    elif button_number == 8:
        flow_run()
        result_label.config(text="Data Extraction Complete")
    elif button_number == 9:
        save_flow()
        result_label.config(text="Excel Saved!")

sheet_name = 'Step 10'
sheet1_name = 'Step 15'
sheet2_name = 'Step 41'
sheet3_name = '7WeekWorkflows'
sheet4_name = '8WeekWorkflows'
file_path = r'C:\Users\GADZINDA\OneDrive - Danone\Desktop\PAT.xlsx'
s10 = pd.read_excel(file_path, sheet_name=sheet_name)
s15 = pd.read_excel(file_path, sheet_name=sheet1_name)
s41 = pd.read_excel(file_path, sheet_name=sheet2_name)
oldwf = pd.read_excel(file_path, sheet_name=sheet3_name)
voldwf = pd.read_excel(file_path, sheet_name=sheet4_name)

# Your variables
num_messages = len(s10['Status'])
num_emails = len(s15['Status'])
num_feedbacks = len(s41['Status'])
num_weekwf = len(oldwf['Status'])
num_vweekwf = len(voldwf['Status'])

def ExcelRead():
    file_path = r'C:\Users\GADZINDA\OneDrive - Danone\Desktop\PAT.xlsx'
    sheet_name = 'Step 10'
    s10 = pd.read_excel(file_path, sheet_name=sheet_name)

def maximize_orsoft_window():
    window_title_pattern = "Orsoft*"  # Match any window title ending with "Excel"
    windows = []
        
    def enum_windows_callback(hwnd, _):
        nonlocal windows
        window_text = win32gui.GetWindowText(hwnd)
        if re.match(window_title_pattern, window_text, re.IGNORECASE):
            windows.append(hwnd)
        
    win32gui.EnumWindows(enum_windows_callback, None)
        
    if windows:
        target_window = windows[0]
        win32gui.ShowWindow(target_window, win32con.SW_RESTORE)  # Restore window if minimized
        win32gui.SetForegroundWindow(target_window)  # Bring to foreground
        win32gui.ShowWindow(target_window, win32con.SW_MAXIMIZE)  # Maximize the window

def save_flow():
# Define the path to the existing Excel file
    existing_file_path = 'C:/Users/GADZINDA/OneDrive - Danone/Desktop/PAT.xlsx'

# Load the existing Excel file
    workbook = openpyxl.load_workbook(existing_file_path)

# Generate today's date in a specific format (e.g., YYYY-MM-DD)
    today_date = datetime.now().strftime("%d.%m.%Y")

# Define the new file name with today's date
    new_file_name = f'PAT_{today_date}.xlsx'

# Define the new path for the saved file, including the new file name
    new_file_path = 'C:/Users/GADZINDA/Danone/Finance Masterdata team - Documents/GL_PROJECT/2. GL Other/4. GL Orsoft Workflow Tracker/' + new_file_name

# Save the workbook to the new location with the new name
    workbook.save(new_file_path)

# Close the workbook
    workbook.close()

def custom_x10():
    try:
        # Call the x10() function here
        ExcelRead()
    except PermissionError:
        result_label.config(text="Close and Save Excel File Before Running Flow!")


def perform_actions(dataframe, column_name):
    for value in dataframe[column_name]:
        pyautogui.typewrite(str(value))
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.leftClick(524, 220)
        
        for _ in range(1):
            pyautogui.press('right')
            time.sleep(0.5)
        
        time.sleep(1)
        pyautogui.leftClick(526, 223)
        
        for _ in range(3):
            pyautogui.press('down')
            time.sleep(0.5)
        
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.leftClick(52, 222)
        time.sleep(1)
        pyautogui.press('left')
        time.sleep(1)
        pyautogui.doubleClick(75, 201)
        time.sleep(1)
        pyautogui.press('backspace')

def step10():
    window.iconify()
    s10 = pd.read_excel(file_path, sheet_name=sheet_name)
    time.sleep(2)
    maximize_orsoft_window()
    time.sleep(1)
    pyautogui.leftClick(48, 130)
    time.sleep(1)
    pyautogui.doubleClick(120, 203)
    time.sleep(1)
    perform_actions(s10, 'WorkFlow ID')
    pyautogui.click(183, 188)


def step15():
    s10 = pd.read_excel(file_path, sheet_name=sheet_name)
    # URL to open
    url = 'https://make.powerautomate.com/environments/Default-4720ed5e-c545-46eb-99a5-958dd333e9f2/flows/cf1ca766-ade4-4833-81cf-4526250a26bb/details'
    webbrowser.open(url)

def step41():
    s10 = pd.read_excel(file_path, sheet_name=sheet_name)
    # URL to open
    url = 'https://make.powerautomate.com/environments/Default-4720ed5e-c545-46eb-99a5-958dd333e9f2/flows/6de0828c-7786-4064-a584-d7d14d80afb3/details'
    webbrowser.open(url)

def step7wwf():
    s10 = pd.read_excel(file_path, sheet_name=sheet_name)
    # URL to open
    url = 'https://make.powerautomate.com/environments/Default-4720ed5e-c545-46eb-99a5-958dd333e9f2/flows/79b29fa3-595a-4469-bd05-66346f48aa7f/details'
    webbrowser.open(url)


def step8wwf():
    window.iconify()
    voldwf = pd.read_excel(file_path, sheet_name=sheet4_name)
    time.sleep(1)
    maximize_orsoft_window()
    time.sleep(1)
    pyautogui.leftClick(48, 130)
    time.sleep(1)
    pyautogui.doubleClick(120, 203)
    time.sleep(1)
    perform_actions(voldwf, 'WorkFlow ID')
    pyautogui.click(183, 188)


def open_excel_file():
    excel_file_path = r"C:\Users\GADZINDA\OneDrive - Danone\Desktop\PAT.xlsx"
    
    # Check if the file exists before attempting to open it
    if os.path.isfile(excel_file_path):
        subprocess.Popen(["start", "excel", excel_file_path], shell=True)
    else:
        result_label.config(text="Excel file not found!")

def refresh_data():
    global num_messages, num_emails, num_feedbacks, num_weekwf, num_vweekwf
    
    # Replace these lines with code to read your Excel data
    s10 = pd.read_excel(file_path, sheet_name=sheet_name)
    s15 = pd.read_excel(file_path, sheet_name=sheet1_name)
    s41 = pd.read_excel(file_path, sheet_name=sheet2_name)
    oldwf = pd.read_excel(file_path, sheet_name=sheet3_name)
    voldwf = pd.read_excel(file_path, sheet_name=sheet4_name)
    # For demonstration purposes, we're using dummy data here
    num_messages = len(s10['Status'])
    num_emails = len(s15['Status'])
    num_feedbacks = len(s41['Status'])
    num_weekwf = len(oldwf['Status'])
    num_vweekwf = len(voldwf['Status'])
    
    # Update the labels to display the refreshed data
    message_label.config(text=f"{num_messages}")
    email_label.config(text=f"{num_emails}")
    feedback_label.config(text=f"{num_feedbacks}")
    weekwf_label.config(text=f"{num_weekwf}")
    vweekwf_label.config(text=f"{num_vweekwf}")

#Run ORSOFT and Download ORS.xlsx
def flow_run():
    def click_ui_element(x, y):
        try:
            pyautogui.click(x, y)
        except Exception as e:
            print(f"Error clicking the UI element: {e}")

    if __name__ == "__main__":
        # Click Windows
        ui_element1_coordinates = (36, 1061)

        # Run Orsoft
        ui_element2_coordinates = (532, 584)

        # Orsoft window maximize
        ui_element3_coordinates = (1620, 59)

        # Connect
        ui_element4_coordinates = (619, 998)

        # Workflows Tab
        ui_element5_coordinates = (83, 125)

        # Excel Open
        ui_element6_coordinates = (244, 79)

        window_title = 'ORSOFT Logon'

        #click_ui_element(*ui_element1_coordinates)

        # Introduce a delay (in seconds) between the clicks (adjust this time as needed)
        time.sleep(1)
        os.startfile("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\ORSOFT User Interface Client 7.5.0 (64-bit)\ORSOFT 7.5.0-GLOBAL.lnk")
        time.sleep(3)

        # Bring the ORSOFT window to the foreground
        orsoft_window = gw.getWindowsWithTitle(window_title)
        if len(orsoft_window) > 0:
            orsoft_window[0].activate()

        time.sleep(2)
        # Maximize the Orsoft window using keyboard shortcut
        pyautogui.hotkey('alt', 'space')
        pyautogui.press('x')


        # Introduce a delay (in seconds) between the clicks (adjust this time as needed)
        time.sleep(2)

        click_ui_element(*ui_element4_coordinates)

    time.sleep(5)

    def main2():

        # Display a message box with two options
        choices = ["Continue", "Interrupt"]
        choice = easygui.buttonbox("Select Continue once the data is loaded", "Data Loading", choices=choices)
            
        if choice == "Continue":
            print("Execution aborted.")
                
        elif choice == "Interrupt":
            print("Script execution interrupted.")
            sys.exit()

    #if __name__ == "__main__":
    #    main2()

    icon_to_click = "ORS-Load"

    r = None
    while r is None:
        r = pyautogui.locateOnScreen('ORS-Load.png', grayscale = True, confidence = 0.8)
    print(icon_to_click + ' now loaded')

    time.sleep(3)
    click_ui_element(*ui_element5_coordinates)
    time.sleep(2)
    click_ui_element(*ui_element6_coordinates)
    time.sleep(5)


    time.sleep(5)
    def maximize_excel_window():
        window_title_pattern = ".*Excel$"  # Match any window title ending with "Excel"
        windows = []
        
        def enum_windows_callback(hwnd, _):
            nonlocal windows
            window_text = win32gui.GetWindowText(hwnd)
            if re.match(window_title_pattern, window_text, re.IGNORECASE):
                windows.append(hwnd)
        
        win32gui.EnumWindows(enum_windows_callback, None)
        
        if windows:
            target_window = windows[0]
            win32gui.ShowWindow(target_window, win32con.SW_RESTORE)  # Restore window if minimized
            win32gui.SetForegroundWindow(target_window)  # Bring to foreground
            win32gui.ShowWindow(target_window, win32con.SW_MAXIMIZE)  # Maximize the window

    if __name__ == "__main__":
        maximize_excel_window()


    time.sleep(3)

    pyautogui.click(46, 76)
    time.sleep(1)
    pyautogui.click(109, 478)
    time.sleep(1)
    pyautogui.click(1078, 188)
    time.sleep(1)
    pyautogui.typewrite("ORS470")
    time.sleep(1)
    #Click Scroll Up 3x
    pyautogui.click(892, 240)
    pyautogui.click(892, 240)
    pyautogui.click(892, 240)
    #Click Desktop
    time.sleep(1)
    pyautogui.click(803, 288)
    #Click Save
    time.sleep(1)
    pyautogui.click(1366, 755)
    #Confirm overwrite
    time.sleep(1)
    pyautogui.click(1036, 532)


    time.sleep(3)

    def open_excel_file(file_path):
        try:
            # Create a COM object for Excel application
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = True

            # Open the Excel file
            workbook = excel.Workbooks.Open(file_path)

        except Exception as e:
            print("Error:", e)

    if __name__ == "__main__":
        # Replace 'file_path' with the full path of the Excel file you want to open
        file_path = r'C:\Users\GADZINDA\OneDrive - Danone\Desktop\PAT.xlsx'
        
        open_excel_file(file_path)




    def bring_excel_window_to_foreground():
        # Replace 'Microsoft Excel' with the title of your Excel window
        window_title = 'PAT - Excel'

        # Bring the Excel window to the foreground
        excel_window = gw.getWindowsWithTitle(window_title)
        if len(excel_window) > 0:
            excel_window[0].activate()

        # Maximize the Excel window using keyboard shortcut
        pyautogui.hotkey('alt', 'space')
        pyautogui.press('x')

    if __name__ == "__main__":
        bring_excel_window_to_foreground()
        time.sleep(3)  # Add a delay to ensure the window is maximized before continuing with other actions
        print("Excel window brought to foreground and maximized.")

    #Data Refresh - PAT.xlsx UI Scripting 

    time.sleep(3)
    pyautogui.click(571, 75)
    time.sleep(1)
    pyautogui.click(525, 127)
    time.sleep(10)

    icon_to_click2 = "PAT-Saved"

    r = None
    while r is None:
        r = pyautogui.locateOnScreen('PAT-Saved.png', grayscale = True, confidence = 0.8)
    print(icon_to_click2 + ' now loaded')

    time.sleep(5)
    pyautogui.click(1884, 22)
    time.sleep(2)
    #pyautogui.press('s')
    time.sleep(2)
    window.attributes("-topmost", True)


# Your variables
num_messages = len(s10['Status'])
num_emails = len(s15['Status'])
num_feedbacks = len(s41['Status'])
num_weekwf = len(oldwf['Status'])
num_vweekwf = len(voldwf['Status'])

# Create the main window
window = tk.Tk()
window.title("PAT")
window.iconbitmap("C:\\Users\\GADZINDA\\Desktop\\PAT\\icon.ico")
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
window_width = 1280
window_height = 680
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
window.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

# Create a frame to hold the big title and logo
title_logo_frame = ttk.Frame(window, padding=20)
title_logo_frame.pack()

# Create a label for the big title
big_title_label = ttk.Label(title_logo_frame, text="PAT Software", font=('Helvetica', 24))
big_title_label.pack()

# Load the logo image
logo_image = Image.open("Bolt.png")  # Replace with the actual path to your logo image
logo_image = logo_image.resize((150, 150))  # Resize the image as needed
logo_photo = ImageTk.PhotoImage(logo_image)

# Create a label to display the logo
logo_label = ttk.Label(title_logo_frame, image=logo_photo)
logo_label.image = logo_photo  # Keep a reference to the image to prevent it from being garbage collected
logo_label.pack()

# Create a frame to hold the buttons and labels
frame = ttk.Frame(window, padding=20)
frame.pack()

# Create styled buttons with larger labels under them
font = ('Helvetica', 18)  # Font size is 18

message_button = ttk.Button(frame, text="Step 10", command=lambda: button_click(1), width=15)
message_button.grid(row=5, column=0, padx=20, pady=10)

email_button = ttk.Button(frame, text="Step 15", command=lambda: button_click(2), width=15)
email_button.grid(row=5, column=1, padx=20, pady=10)

feedback_button = ttk.Button(frame, text="Step 41", command=lambda: button_click(3), width=15)
feedback_button.grid(row=5, column=2, padx=20, pady=10)

weekwf_button = ttk.Button(frame, text="7 Week WF", command=lambda: button_click(4), width=15)
weekwf_button.grid(row=5, column=3, padx=20, pady=10)

vweekwf_button = ttk.Button(frame, text="8 Week WF", command=lambda: button_click(5), width=15)
vweekwf_button.grid(row=5, column=4, padx=20, pady=10)

# Create labels with larger centered numbers
message_label = ttk.Label(frame, text=f"{num_messages}", font=font)
message_label.grid(row=4, column=0)

email_label = ttk.Label(frame, text=f"{num_emails}", font=font)
email_label.grid(row=4, column=1)

feedback_label = ttk.Label(frame, text=f"{num_feedbacks}", font=font)
feedback_label.grid(row=4, column=2)

weekwf_label = ttk.Label(frame, text=f"{num_weekwf}", font=font)
weekwf_label.grid(row=4, column=3)

vweekwf_label = ttk.Label(frame, text=f"{num_vweekwf}", font=font)
vweekwf_label.grid(row=4, column=4)

# Create a big button at the center of the bottom part of the window
big_button = ttk.Button(frame, text="Refresh Data", command=lambda: button_click(6), width=30)
big_button.grid(row=10, column=0, columnspan=5, pady=20)

# Create an "Open Excel" button of the same size below the "Big Button"
excel_button = ttk.Button(frame, text="Open Excel", command=lambda: button_click(7), width=30)
excel_button.grid(row=11, column=0, columnspan=5, pady=20)

# Create an "Run Flow" button of the same size below the "Excel Button"
flow_button = ttk.Button(frame, text="Run Extraction Flow", command=lambda: button_click(8), width=30)
flow_button.grid(row=12, column=0, columnspan=5, pady=20)

# Create an "Run Flow" button of the same size below the "Excel Button"
flow_button = ttk.Button(frame, text="Save Excel", command=lambda: button_click(9), width=30)
flow_button.grid(row=13, column=0, columnspan=5, pady=20)

# Create a label to display the result
result_label = ttk.Label(window, text="", font=font)
result_label.pack()

# Start the Tkinter main loop
window.mainloop()
