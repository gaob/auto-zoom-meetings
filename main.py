import time, datetime, platform
import getpass
import subprocess
import pyautogui
from pathlib import Path
from datetime import date

current_os = platform.system()


def launch_zoom(path_home):
    zoom = get_zoom(pyautogui.getWindowsWithTitle("Zoom"))
    if (zoom):
        print("Obtained Zoom: "+str(zoom.title))
        zoom.activate()
        return zoom
    else:
        zoom_path = path_home+r'\AppData\Roaming\Zoom\bin\Zoom.exe'
        return subprocess.Popen(zoom_path, stdout=subprocess.PIPE)


def get_zoom(zoom_list):
    for zoom in zoom_list:
        if (zoom.title=="Zoom" or zoom.title=="Zoom Cloud Meetings"):
            return zoom

    return None


def main():
    path_home=str(Path.home())
    print("User Home Directory: "+path_home)
    print("Zoom Email:")
    user_email = input()
    user_password = getpass.getpass()
    print("Name:")
    user_name = input()
    print("Getting courses...")
    courses = get_courses("courses.txt")  # [[course1, 15:40], [course2, 08:40]]
    print("Done!")
    print("Please wait. You will be logged in automatically when the time comes. Do not close this window. [Press CTRL + C to terminate!]")
    while True:
        all_done = True
        for course in courses:
            all_done = True
            if not course[3] and is_today(course[5]):  # Not marked as done yet.
                all_done = False
                if not course_finished(course[1], course[2]):  # it's not this course's time yet.
                    continue
                else:   # it's showtime!
                    zoom_automate(course[0], course[4], user_email, user_password, user_name, course[1] + 2, course[2] - 10, path_home)  # 2 x 50 min, 1 x 10 min
                    course[3] = True  # Mark course as done.
            else:   # This course has been attended.
                continue
        if all_done:
            break
        time.sleep(60)
    exit(0)
    

def zoom_automate(zoom_id, course_password, user_email, user_password, user_name, term_hour, term_minute, path_home):
    zoom = launch_zoom(path_home)
    print("Wait for zoom to launch.")
    time.sleep(5)
    fw = pyautogui.getActiveWindow()
    if (fw.title=="Zoom Cloud Meetings"):
        logged_in = False
    else:
        logged_in = True

    if not logged_in:
        sign_in(zoom, user_email, user_password)
        print("you are logged in.")
    else:
        print("you are already logged in.")

    join_meeting(zoom, zoom_id, user_name, course_password)

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
    fw = pyautogui.getActiveWindow()
    pyautogui.click(fw.left+400, fw.top+335)
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


def course_finished(hour, minute):
    if hour > 23 or hour < 0:
        hour %= 24
    if minute > 59 or minute < 0:
        minute %= 60

    now = datetime.datetime.now()
    course_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    five_minutes = datetime.timedelta(minutes=5)
    if now >= (course_time - five_minutes):
        return True
    print("next course joins at: "+str(course_time - five_minutes)+"\tcurrent time: "+str(now))
    return False


def is_today(weekday):
    if weekday == -1:
        return True
    if weekday == date.today().weekday():
        return True
    return False


def get_courses(path):
    wrapper = []
    try:
        fp = open('courses.txt', 'r')
        courses = fp.readlines()
        for course in courses:
            list = course.split("/")
            zoom_id = list[0].strip().replace("-", "")
            course_time = list[1].strip().split(":")
            if (len(list) > 2):
                zoom_password = list[2].strip()
            else:
                zoom_password = ''
            if (len(list) > 3):
                weekday = int(list[3].strip())
            else:
                weekday = -1
            wrapper.append([zoom_id, int(course_time[0]), int(course_time[1]), False, zoom_password, weekday])
    finally:
        fp.close()
    return wrapper


def zoom_automate2(zoom_id, user_email, user_password, term_hour, term_minute):
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
    print("pyautogui.click()")
    # pyautogui.click(fw.left+308, fw.top+240)


if __name__ == '__main__':
    main()
