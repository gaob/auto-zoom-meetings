from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import selenium.common.exceptions
import time, datetime, platform
import getpass
import subprocess
import pyautogui

current_os = platform.system()


def launch_zoom():
    return subprocess.Popen(r'C:\Users\likes\AppData\Roaming\Zoom\bin\Zoom.exe', stdout=subprocess.PIPE)


def main():
    print(pyautogui.position())
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
            if not course[3]:  # Not marked as done yet.
                all_done = False
                if not course_finished(course[1], course[2]):  # it's not this course's time yet.
                    continue
                else:   # it's showtime!
                    zoom_automate(course[0], user_email, user_password, user_name, course[1] + 2, course[2] - 10)  # 2 x 50 min, 1 x 10 min
                    course[3] = True  # Mark course as done.
            else:   # This course has been attended.
                continue
        if all_done:
            break
        time.sleep(30)
    exit(0)


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
    

def zoom_automate(zoom_id, user_email, user_password, user_name, term_hour, term_minute):
    zoom = launch_zoom()
    print("Wait for zoom to launch.")
    time.sleep(5)
    fw = pyautogui.getActiveWindow()
    if (fw.height<=400):
        logged_in = False
    else:
        logged_in = True

    if not logged_in:
        sign_in(zoom, user_email, user_password)
        print("you are logged in.")
    else:
        print("you are already logged in.")

    join_meeting(zoom, zoom_id, user_name)

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


def join_meeting(zoom, meeting_number, user_name, meeting_password = ''):
    print("Begin to join meeting: "+str(meeting_number))
    fw = pyautogui.getActiveWindow()
    pyautogui.click(fw.left+400, fw.top+335)
    time.sleep(1)
    pyautogui.write(meeting_number)
    pyautogui.write(['tab', 'tab'])
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(['backspace'])
    time.sleep(1)
    pyautogui.write(user_name)


def join_meeting2(browser, meeting_number):
    logged_join_button = browser.find_element_by_xpath('//*[@id="btnJoinMeeting"]')
    logged_join_button.click()
    browser.find_element_by_xpath('//*[@id="join-confno"]').send_keys(str(meeting_number[0:3]))
    browser.implicitly_wait(3)  # seconds
    browser.find_element_by_xpath('//*[@id="join-confno"]').send_keys(str(meeting_number[3:6]))
    browser.implicitly_wait(3)  # seconds
    browser.find_element_by_xpath('//*[@id="join-confno"]').send_keys(str(meeting_number[6:]))
    browser.find_element_by_xpath('//*[@id="btnSubmit"]').click()
    time.sleep(2)
    try:
        browser.find_element_by_xpath('//*[@id="action_container"]/div[3]/a').click()
    except Exception:
        try:
            browser.find_element_by_xpath('//*[@id="launch_meeting"]/div/div[4]/a').click()
        except Exception:
            print("Couldn't find the WC link. Moving on.")

    zoom_root_url = browser.current_url.split("//")[-1].split("/")[0]
    destination_url = zoom_root_url + "/wc/join/" + meeting_number + "?pwd="
    browser.get("https://" + destination_url)
    browser.maximize_window()
    try:
        browser.find_element_by_xpath('//*[@id="joinBtn"]').click()
    except Exception as e:
        print("Could not click on button. Moving on.")
        print(e)
    time.sleep(3)
    try:
        browser.find_element_by_xpath('//*[@id="dialog-join"]/div[4]/div/div/div[1]/button').click()
    except Exception as e:
        print("Could not click on button. Moving on.")
        print(e)


def course_finished(hour, minute):
    if hour > 23 or hour < 0:
        hour %= 24
    if minute > 59 or minute < 0:
        minute %= 60

    now = datetime.datetime.now()
    course_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if now >= course_time:
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
            wrapper.append([zoom_id, int(course_time[0]), int(course_time[1]), False])
    finally:
        fp.close()
    return wrapper


if __name__ == '__main__':
    main()
