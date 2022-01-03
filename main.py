# Libraries
import win32api
import win32con
import win32gui

import pyautogui
import keyboard

import msvcrt
import time
import os
import ctypes
import termcolor

# global variables
windows = []
windowsHwnd = []
selectedWindowIndex = 0

intervalTime = None
mouseX = None
mouseY = None


# Lists the open windows
# for user to select
def list_window_names():
    global windows

    def winEnumHandler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            windowName = win32gui.GetWindowText(hwnd)
            if windowName != "":
                windows.append(hex(hwnd) + " " + windowName)
                windowsHwnd.append(hwnd)

    win32gui.EnumWindows(winEnumHandler, None)


# Moves cursor
def move_cursor():
    refresh_print()

    # Removes extra cursors and reset colors
    for i in range(len(windows)):
        windows[i] = windows[i].replace("> ", "")
        windows[i] = termcolor.colored(windows[i], "white")

    # Adds cursor to selected window
    windows[selectedWindowIndex] = termcolor.colored("> ", "green") + windows[selectedWindowIndex]

    # output
    for i in windows:
        print(i + " " * 10)  # extra space to overwrite previous cursor


# Overwrites the output
def refresh_print():
    backChar = "\033[F"
    print(backChar * (len(windows) + 1))


# Selects the window
def select_window():
    global selectedWindowIndex

    # print options first
    for i in windows:
        print(i)

    move_cursor()
    while True:
        key = ord(msvcrt.getch())
        if key == 13:  # Enter is pressed
            break
        elif key == 224:
            key = ord(msvcrt.getch())
            if key == 72:  # Up arrow is pressed
                if selectedWindowIndex < 0:
                    selectedWindowIndex = len(windows) - 1
                else:
                    selectedWindowIndex -= 1
                move_cursor()
            elif key == 80:  # Down arrow is pressed
                if selectedWindowIndex >= len(windows) - 1:
                    selectedWindowIndex = 0
                else:
                    selectedWindowIndex += 1
                move_cursor()
        else:
            pass


# Function to get the time interval between clicks
def get_interval_time():
    global intervalTime
    print(termcolor.colored("\nEnter the time interval between clicks (in seconds): ", "yellow"))
    intervalTime = input()

    # Check if the input is a number
    while True:
        try:
            float(intervalTime)
            break
        except ValueError:
            print(termcolor.colored("Invalid Input!", "red"))
            print(termcolor.colored("Enter the time interval between clicks (in seconds): ", "yellow"))
            intervalTime = input()

    # cast intervalTime to float
    intervalTime = float(intervalTime)


# Gets user's mouse position
def get_mouse_position():
    global mouseX, mouseY

    print(termcolor.colored("Hover your mouse over the position you want to click,", "yellow"))
    print(termcolor.colored("then press F6 to continue.", "yellow"))

    while True:
        if keyboard.is_pressed("f6"):
            # Set mouse position
            mouseX, mouseY = pyautogui.position()
            break


# Debug Logging function
def debug_log(message):
    print(termcolor.colored("[DEBUG] ", "yellow") + message)


# main func
def main():
    # variables
    isAutoClick = False

    # prime cmd for color
    kernel32 = ctypes.WinDLL('kernel32')
    hStdOut = kernel32.GetStdHandle(-11)
    mode = ctypes.c_ulong()
    kernel32.GetConsoleMode(hStdOut, ctypes.byref(mode))
    mode.value |= 4
    kernel32.SetConsoleMode(hStdOut, mode)

    # main loop
    output = """
========== Auto Clicker ==========
Select the window that you would like to click on (using up, down and enter):
"""
    print(termcolor.colored(output, "yellow"))
    list_window_names()
    select_window()

    # get interval time
    get_interval_time()

    # Get mouse position to click
    get_mouse_position()

    print(termcolor.colored("\nPress F6 to start/stop clicking.", "yellow"))
    print(termcolor.colored("Press F7 to exit.", "yellow"))

    # Cache hwnd
    hwnd = windowsHwnd[selectedWindowIndex]

    # Click on selected window when F6 is pressed
    # TODO: Add threading to auto clicking
    while True:
        if keyboard.is_pressed("f6"):
            isAutoClick = not isAutoClick

        if keyboard.is_pressed("f7"):
            break

        if isAutoClick:
            # TODO: Fix this
            # Cast x and y to LPARAM
            downLParam = win32api.MAKELONG(win32con.HTCLIENT, win32con.WM_LBUTTONDOWN)
            upLParam = win32api.MAKELONG(win32con.HTCLIENT, win32con.WM_LBUTTONUP)
            mouseLParam = win32api.MAKELONG(mouseX, mouseY)

            # Click on selected window
            win32api.SendMessage(hwnd, win32con.WM_SETCURSOR, hwnd, downLParam)
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, mouseLParam)

            win32api.SendMessage(hwnd, win32con.WM_SETCURSOR, hwnd, upLParam)
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, None, mouseLParam)

            debug_log("Clicked on window: " + windows[selectedWindowIndex])

            time.sleep(intervalTime)


if __name__ == "__main__":
    main()
    os.system("pause")
    os.system("cls")

    # TODO: Handle program fault
