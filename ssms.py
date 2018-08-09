"""
Purpose:Connect to SQL Server using Enterprise Manager with Windows Authentication
Date                Programmer              Description
July 20, 2018       Yury Stanev             Original
TODO: convert to '.exe' using 'pyinstaller -F ssms.py' on Windows
"""

import sqlite3  # creates connection to the DB file
from prettytable import from_db_cursor
import getpass
import subprocess  # allows you to spawn new processes, connect to their input/output/error pipes
import time
import pyautogui

conn = sqlite3.connect('users2.sqlite')
c = conn.cursor()  # allows to perform SQL operations


# Functions
def get_app_id():
    # list all the applications
    c.execute("SELECT DISTINCT * FROM apps ORDER BY appid ASC")
    table = from_db_cursor(c)
    print(table)

    app_id = input("Pick app code: ")
    while app_id == "":
        app_id = input("Pick app code: ")

    return ''.join(app_id)


def get_group_id(app_id):
    # list the group IDs associated with the application
    c.execute("SELECT DISTINCT id, name FROM groups WHERE appid = '%s'" % app_id)
    table = from_db_cursor(c)
    print(table)

    group_id = input("Pick group ID: ")
    while group_id == "":
        group_id = input("Pick group ID: ")

    c.execute("SELECT DISTINCT name FROM groups WHERE id = '%s'" % group_id)
    group_id = c.fetchone()

    pswd = getpass.getpass()  # have to use 'terminal', default prompt: 'Password: '

    return ''.join(group_id), pswd  # ''.join() is used to convert a tuple to a string


def get_instance(app_id):
    # list the instances associated with the application
    c.execute("SELECT DISTINCT id, name FROM instances WHERE appid = '%s'" % app_id)
    table = from_db_cursor(c)
    print(table)

    instance_name = input("Pick instance name: ")
    while instance_name == "":
        instance_name = input("Pick instance name: ")

    c.execute("SELECT DISTINCT name FROM instances WHERE id = '%s'" % instance_name)
    instance_name = c.fetchone()

    return ''.join(instance_name)  # ''.join() is used to convert a tuple to a string


def pick_sql_version():  # might be redundant

    # create an easy list to pick from
    sql_year = ['2012', '2014', '2016']
    for i, val in enumerate(sql_year, start=1):
        print(i, " SQL Server Management Studio ", val)

    sql_version = input("Pick the SSMS version: ")
    while sql_version == "":
        sql_version = input("Pick the SSMS version: ")

    # depending on the sql version path might vary
    version_d = {
        # TODO: add path for all SSMS versions - DONE
        '1': '110',
        '2': '120',
        '3': '130'
    }

    # If invalid path is chosen fall back to default (120)
    path = '"C:/Program Files (x86)/Microsoft SQL Server/{}/Tools/Binn/ManagementStudio/Ssms.exe"'.format(
        version_d.get(sql_version, '120'))

    return path


def command_builder(path, group_id):  # builds a command used to launch SSMS
    command = 'runas /netonly /user:' + group_id + " " + path
    return command


def start_ssms(command, pswd, instance_name):

    # subprocess.Popen(['C:/Windows/System32/cmd.exe'], stdin=subprocess.DEVNULL)  # 'opens cmd.exe'
    subprocess.Popen(['start', '/wait', 'cmd'], shell=True)
    time.sleep(1)

    pyautogui.typewrite(command)  # enters the command to start SSMS
    pyautogui.press('enter')
    pyautogui.typewrite(pswd)  # enter password
    pyautogui.press('enter')

    # print (instance_name)

    time.sleep(35)  # waiting for SSMS to start

    # because SSMS defaults to entering instance name after start
    # pyautogui.press('backspace')
    pyautogui.typewrite(instance_name)
    pyautogui.press('enter')  # press 'enter' to connect to instance

    # click the 'Connect' button
    #  TODO: Get 'connect' button coordinates on 1920x1080 screen - DONE?
    #  TODO: Adjust screenshots for automation
    # connect_button = pyautogui.locateOnScreen('imgs/connect_bt.png', grayscale=True, region=(856, 556, 82, 33))
    # connect_x, connect_y = pyautogui.center(connect_button)
    # pyautogui.click(connect_x, connect_y, duration=0.75)


# Functions Calls
def main():
    app_id = get_app_id()  # stores returned value, allows values to be passed around and accessed later
    instance = get_instance(app_id)
    path = pick_sql_version()
    group_id, pswd = get_group_id(app_id)
    command = command_builder(path, group_id)
    start_ssms(command, pswd, instance)


main()