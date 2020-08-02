import time, datetime, platform
import getpass
import subprocess
import pyautogui
import webbrowser
import pickle
import os.path
from pathlib import Path
from datetime import date
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

current_os = platform.system()

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1TgnLNDqjqH_TNAyxotZNgICu4WOLcRwDlpN0O6Jx1lc'
SAMPLE_RANGE_NAME = 'Courses'


def launch_zoom(path_home):
    zoom = get_zoom(pyautogui.getWindowsWithTitle("Zoom"))
    if (zoom):
        zoom.activate()
        return zoom
    else:
        zoom_path = path_home+r'\AppData\Roaming\Zoom\bin\Zoom.exe'
        return subprocess.Popen(zoom_path, stdout=subprocess.PIPE)


def launch_meet(meet_id):
    webbrowser.open('http://meet.google.com/'+meet_id)
    time.sleep(5)
    meet = get_meet()
    meet.activate()
    meet.maximize()
    time.sleep(5)
    return meet


def get_meet():
    meet_list = pyautogui.getWindowsWithTitle("Meet")
    for meet in meet_list:
        if (meet.title.startswith("Meet")):
            print("Obtained Meet: "+str(meet.title))
            return meet


def get_zoom(zoom_list):
    for zoom in zoom_list:
        if (zoom.title=="Zoom" or zoom.title=="Zoom Cloud Meetings"):
            print("Obtained Zoom: "+str(zoom.title))
            return zoom

    return None


def main():
    attended = {}
    
    path_home=str(Path.home())
    print("User Home Directory: "+path_home)
    user_info = get_zoom_userinfo_from(path_home)
    user_email = None
    user_password = None
    user_name = None
    if user_info:
        user_email = user_info[0]
        user_password = user_info[1]
        user_name = user_info[2]
    else:
        print("Zoom Email:")
        user_email = input()
        user_password = getpass.getpass()
        print("Name:")
        user_name = input()
    if (not user_name):
        if (not user_email):
            user_name = "student"
        else:
            user_name = user_email
    print("Please wait. You will be logged in automatically when the time comes. Do not close this window. [Press CTRL + C to terminate!]")
    while True:
        while True:
            courses = get_courses()
            print("Refreshed courses...")
            all_done = True
            for course in courses:
                all_done = True
                if is_today(course[5]) and (course[0]+str(course[1])+str(course[2])) not in attended:  # Not marked as done yet.
                    all_done = False
                    if not course_approaching(course[0], course[1], course[2]):  # it's not this course's time yet.
                        continue
                    else:   # it's showtime!
                        course_automate(course[0], course[4], user_email, user_password, user_name, course[1] + 2, course[2] - 10, path_home)  # 2 x 50 min, 1 x 10 min
                        attended[(course[0]+str(course[1])+str(course[2]))] = True  # Mark course as done.
                else:   # This course has been attended.
                    continue
            if all_done:
                break
            time.sleep(60)
        print("No new courses. Have attended: "+str(attended))
        time.sleep(300)
    exit(0)
    

def course_automate(course_id, course_password, user_email, user_password, user_name, term_hour, term_minute, path_home):
    if course_id.isdigit():
        zoom_automate(course_id, course_password, user_email, user_password, user_name, path_home)
    else:
        meet_automate(course_id)

    return

    while not course_finished(term_hour, term_minute):
        time.sleep(60)
    # Exit buttons
    try:
        browser.find_element_by_xpath('//*[@id="wc-footer"]/div[3]/button').click()  # Are you sure button
        browser.find_element_by_xpath('/html/body/div[11]/div/div/div/div[2]/div/div/button').click()  # Yes button.
    except Exception as e:
        print(e)
        print("Couldn't exit gracefully.")
    time.sleep(200)
    browser.quit()        


def meet_automate(meet_id):
    meet = launch_meet(meet_id)
    pyautogui.click(meet.left+700, meet.top+750)
    pyautogui.click(meet.left+1300, meet.top+600)


def zoom_automate(zoom_id, course_password, user_email, user_password, user_name, path_home):
    zoom = launch_zoom(path_home)
    print("Wait for zoom to launch.")
    time.sleep(5)
    zoom = pyautogui.getActiveWindow()
    if (zoom.title=="Zoom Cloud Meetings"):
        logged_in = False
    else:
        logged_in = True

    if not logged_in:
        sign_in(zoom, user_email, user_password)
        print("you are logged in.")
    else:
        print("you are already logged in.")

    join_meeting(zoom, zoom_id, user_name, course_password)


def sign_in(zoom, user_email, user_password):
    pyautogui.write(['tab', 'tab', 'enter'])
    time.sleep(1)
    pyautogui.write(user_email)
    time.sleep(1)
    pyautogui.write(['tab'])
    time.sleep(1)
    pyautogui.write(user_password)
    time.sleep(1)
    pyautogui.write(['enter'])
    print("Wait for zoom to log in.")
    time.sleep(5)


def join_meeting(zoom, meeting_number, user_name, meeting_password):
    print("Begin to join meeting: "+str(meeting_number))
    zoom.maximize()
    time.sleep(2)
    pyautogui.click(zoom.left+830, zoom.top+450)
    time.sleep(2)
    pyautogui.write(meeting_number)
    pyautogui.write(['tab', 'tab'])
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(['backspace'])
    time.sleep(1)
    pyautogui.write(user_name)
    pyautogui.write(['enter'])
    time.sleep(1)
    pyautogui.write(meeting_password)
    time.sleep(1)
    pyautogui.write(['enter'])
    time.sleep(1)
    pyautogui.write(['enter'])    


def course_approaching(course_id, hour, minute):
    if hour > 23 or hour < 0:
        hour %= 24
    if minute > 59 or minute < 0:
        minute %= 60

    now = datetime.datetime.now()
    course_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    five_minutes = datetime.timedelta(minutes=5)
    if now >= (course_time - five_minutes):
        return True
    print("next course " + course_id + " joins at: "+str(course_time - five_minutes)+"\tcurrent time: "+str(now))
    return False


def is_today(weekday):
    if weekday == -1:
        return True
    if weekday == date.today().weekday():
        return True
    return False


def get_zoom_userinfo_from(path):
    wrapper = None
    try:
        fp = open('userinfo.txt', 'r')
        user_email = fp.readline()
        user_password = fp.readline()
        user_name = fp.readline()

        wrapper = [user_email, user_password, user_name]
    finally:
        fp.close()
    return wrapper


def get_courses():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    wrapper = []
    if not values:
        print('No data found.')
    else:
        for row in values:
            zoom_id = row[0].strip().replace("-", "")
            course_time = row[1].strip().split(":")
            if (len(row)>2):
                weekday = int(row[2].strip())
            else:
                weekday = -1
            if (len(row)>3):
                zoom_password = row[3].strip()
            else:
                zoom_password = ''                
            wrapper.append([zoom_id, int(course_time[0]), int(course_time[1]), False, zoom_password, weekday])
    return wrapper            


def pyautogui_code(zoom_id, user_email, user_password, term_hour, term_minute):
    time.sleep(5)
    pyautogui.write(['tab', 'tab', 'enter'])
    time.sleep(1)
    pyautogui.write(user_email)
    time.sleep(1)
    pyautogui.write(['tab'])
    time.sleep(1)
    pyautogui.write(user_password)
    time.sleep(1)
    pyautogui.write(['enter'])
    pyautogui.moveTo(810, 445)
    fw = pyautogui.getActiveWindow()
    print(str(pyautogui.position().x-fw.left)+','+str(pyautogui.position().y-fw.top))
    print(pyautogui.getAllTitles)
    print("pyautogui.click()")
    # pyautogui.click(fw.left+308, fw.top+240)


if __name__ == '__main__':
    main()
