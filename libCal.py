import collections
import json
import os
import platform
import random
import string
import tkMessageBox
from Tkinter import *
from ttk import Progressbar

sys.path.append(os.getcwd() + '/bin')
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import httplib2
import os
import re

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'LibCalBooker'


class GUI:
    def __init__(self, master):
        ################-GLOBAL-VARS-############################
        self.version = 1.1
        self.loadingMsg = "Loading......................."
        self.driver = ""
        self.tupleDates = []
        self.roomTimeList = []
        self.roomIndex = 0
        self.outputTimeArray = []
        self.userInfo = []
        self.master = master
        self.outputEmail = ""

        ###############-WINDOW-SETUP-############################
        self.master.title("LibCal Booker v" + str(self.version))
        self.master.resizable(width=False, height=False)
        self.master.protocol("WM_DELETE_WINDOW", self.window_close)
        self.master.createcommand('exit', self.window_close)
        self.master.bind('<Return>', self.submit_click)

        ###############-LOAD-FILE-###############################
        # TODO: Check fields of json, if do not match, transfer over files
        try:
            with open("userInfo.json") as data_file:
                self.userInfo = json.load(data_file)
        except:
            print "No existing user"
            with open("userInfo.json", "w+") as data_file:
                self.userInfo = dict(version=self.version, first="Mitchell", last="Pynn", email="mpynn@lakeheadu.ca", override=0, confirm=1,
                                     browser=0, firstLoad=True, authEmail="")
                json.dump(self.userInfo, data_file)

        ###################-DATE SELECTION-######################
        # set up availDates
        self.availDates = [self.loadingMsg]
        # set up chosen date and give default value
        self.chosenDate = StringVar(self.master)
        self.chosenDate.set(self.loadingMsg)
        # set up dateOptionMenu
        self.dateOptionMenu = OptionMenu(self.master, self.chosenDate, *self.availDates, command=self.date_click)
        self.dateOptionMenu.grid(row=0, column=0, columnspan=1, sticky=(N, S, E, W), pady=(8, 0), padx=(5, 0))

        ##################-ROOM SELECTION-######################
        # set up availRooms
        self.availRooms = ['LI 1004', 'LI 1006', 'LI 1007', 'LI 1008', 'LI 1009', 'LI 1010', 'LI 4001', 'LI 4002',
                           'LI 4003', 'LI 4004', 'LI 4005', 'LI 4006', 'LI 4007', 'LI 4008', 'LI 4009', 'LI 4010']
        self.availTimes = ['8:00am - 8:30am', '8:30am - 9:00am', '9:00am - 9:30am', '9:30am - 10:00am',
                           '10:00am - 10:30am',
                           '10:30am - 11:00am', '11:00am - 11:30am', '11:30am - 12:00pm', '12:00pm - 12:30pm',
                           '12:30pm - 1:00pm',
                           '1:00pm - 1:30pm', '1:30pm - 2:00pm', '2:00pm - 2:30pm', '2:30pm - 3:00pm',
                           '3:00pm - 3:30pm',
                           '3:30pm - 4:00pm', '4:00pm - 4:30pm', '4:30pm - 5:00pm', '5:00pm - 5:30pm',
                           '5:30pm - 6:00pm', '6:00pm - 6:30pm',
                           '6:30pm - 7:00pm', '7:00pm - 7:30pm', '7:30pm - 8:00pm', '8:00pm - 8:30pm',
                           '8:30pm - 9:00pm', '9:00pm - 9:30pm',
                           '9:30pm - 10:00pm', '10:00pm - 10:30pm', '10:30pm - 11:00pm', '11:00pm - 11:30pm',
                           '11:30pm - 11:59pm']
        # set up chosen room and give default value
        self.chosenRoom = StringVar(self.master)
        self.chosenRoom.set(self.availRooms[0])
        # set up roomOptionMenu
        self.roomOptionMenu = OptionMenu(self.master, self.chosenRoom, *self.availRooms, command=self.room_click)
        self.roomOptionMenu.grid(row=1, column=0, columnspan=1, sticky=(N, S, E, W), padx=(5, 0), pady=(5, 0))

        ##################-TIME SELECTION-####################
        self.timeOptionList = Listbox(self.master, selectmode=EXTENDED, height=20, exportselection=0, takefocus=0)
        self.timeOptionList.grid(row=2, column=0, rowspan=200, columnspan=1, sticky=(N, S, E, W), padx=(5, 0), pady=5)
        self.timeOptionList.insert(0, self.loadingMsg)
        self.timeOptionList.config(state=DISABLED)

        #################-BUTTONS-##########################
        # user info
        self.infoLabel = Label(self.master, text="[ U S E R  I N F O ]", font=("Helvetica", 16, "bold"))
        self.infoLabel.grid(row=0, column=1, columnspan=2, sticky=E + W)

        # first Name label and input
        self.fnameLabel = Label(self.master, text="First: ", font=("Helvetica", 12, "bold"))
        self.fnameLabel.grid(row=1, column=1, sticky=W)

        self.fnameEntry = Entry(self.master)
        self.fnameEntry.grid(row=1, column=2, stick=E + W, padx=(0, 5))
        self.fnameEntry.insert(0, self.userInfo["first"])

        # last name label and input
        self.lnameLabel = Label(self.master, text="Last: ", font=("Helvetica", 12, "bold"))
        self.lnameLabel.grid(row=2, column=1, sticky=W)

        self.lnameEntry = Entry(self.master)
        self.lnameEntry.grid(row=2, column=2, stick=E + W, padx=(0, 5))
        self.lnameEntry.insert(0, self.userInfo["last"])

        # email label and entry
        self.emailLabel = Label(self.master, text="Email: ", font=("Helvetica", 12, "bold"))
        self.emailLabel.grid(row=3, column=1, sticky=W)

        self.emailEntry = Entry(self.master)
        self.emailEntry.grid(row=3, column=2, stick=E + W, padx=(0, 5))
        self.emailEntry.insert(0, self.userInfo["email"])

        # showBrowser checkbox
        self.browserVal = IntVar(self.master)
        self.browserVal.set(self.userInfo["browser"])
        self.browser = Checkbutton(self.master, text="Show web browser", variable=self.browserVal,
                                   command=self.browser_show, onvalue=1, offvalue=0, font=("Helvetica", 12),
                                   takefocus=0)
        self.browser.grid(row=4, column=2, sticky=W)

        # override checkbox
        self.overrideVal = IntVar(self.master)
        self.overrideVal.set(self.userInfo["override"])
        self.override = Checkbutton(self.master, text="Override 2hr max", variable=self.overrideVal,
                                    onvalue=1, offvalue=0, font=("Helvetica", 12), takefocus=0)
        self.override.grid(row=5, column=2, sticky=W)

        # confirm checkbox
        self.confirmVal = IntVar(self.master)
        self.confirmVal.set(self.userInfo["confirm"])
        self.confirm = Checkbutton(self.master, text="Enable confirm dialog", variable=self.confirmVal,
                                   onvalue=1, offvalue=0, font=("Helvetica", 12), takefocus=0)
        self.confirm.grid(row=6, column=2, sticky=W)

        # submit button
        self.submit = Button(self.master, text="Submit", command=self.submit_click, takefocus=0)
        self.submit.grid(row=8, column=2, sticky=(N, S, E, W), padx=(0, 5), pady=(0, 5))
        self.submit["state"] = "disabled"

        # loading bar
        self.loadingBar = Progressbar(self.master, orient=HORIZONTAL, length=100, mode='determinate')
        self.loadingBar.grid(row=9, column=2, sticky=(N, S, E, W), padx=(0, 5), pady=(0, 5))
        self.loadingBar["value"] = 10

        # email
        self.emm = Button(self.master, text="email tester", command=self.test_email_click, takefocus=0)
        self.emm.grid(row=10, column=2, sticky=(N, S, E, W), padx=(0, 5), pady=(0, 5))

        # update skeleton GUI, then load data
        self.master.update()
        self.load_data()

        # make sure window on top
        self.master.lift()

        # show welcome message if first load
        if self.userInfo["firstLoad"]:
            tkMessageBox.showinfo("Welcome",
                                  "Currently, booking multiple rooms is not permitted. You are limited to only booking one room per session. However, booking multiple time slots per room is permitted.\n\nMake sure to update the [ USER INFO ] section with your own name and email. \n\nCreated by Mitchell Pynn ")

    def test_email_click(self):
        self.outputEmail = "mpynn@lakeheadu.ca"

    def email_click(self, randVals):
        """
        Shows basic usage of the Gmail API.

        Creates a Gmail API service object and outputs a list of label names
        of the user's Gmail account.
        """
        credentials = self.get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('gmail', 'v1', http=http)

        userProfile = service.users().getProfile(userId='me').execute()

        # set user JSON auth email
        self.userInfo["authEmail"] = userProfile['emailAddress']

        for idItem in randVals:
            print idItem
            try:
                # please newer_than:1d to: + idItem
                email1 = service.users().messages().list(userId='me',
                                                       q='to:' + idItem).execute()
                output = email1['messages']

                text = service.users().messages().get(userId='me', id=output[0]['id']).execute()
                ampFix = text['snippet'].replace('&amp;', '&')


                confirmUrl = \
                re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+|]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', ampFix)[
                    0] + '&m=confirm'
                print(confirmUrl)

                self.driver.execute_script(
                    "window.open('" + confirmUrl + "', 'new_window')")
                self.driver.switch_to.window(self.driver.window_handles[1])
                try:
                    # wait until page is loaded and success element found
                    element = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "final_msg"))
                    )

                    # close tab and return to main page
                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])

                except:
                    print "Could not book room " + idItem

            except:
                print('An error occurred confirming the room: ????')

    def get_credentials(self):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       self.outputEmail + '.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def browser_show(self):
        if self.browserVal.get() == 1:
            # hide window/throw in corner
            self.driver.set_window_position(0, 0)
            self.driver.set_window_size(100, 600)
        else:
            # hide window/throw in corner
            self.driver.set_window_position(3000, 3000)
            self.driver.set_window_size(100, 300)

    def save_data(self):
        # update userFile when submitted
        with open("userInfo.json", 'r+') as data_file:
            self.userInfo["first"] = self.fnameEntry.get()
            self.userInfo["last"] = self.lnameEntry.get()
            self.userInfo["email"] = self.emailEntry.get()
            self.userInfo["confirm"] = self.confirmVal.get()
            self.userInfo["override"] = self.overrideVal.get()
            self.userInfo["browser"] = self.browserVal.get()
            self.userInfo["firstLoad"] = False

            data_file.seek(0)
            json.dump(self.userInfo, data_file)
            data_file.truncate()

    def window_close(self):
        # save data
        self.save_data()
        # destroy GUI and quit driver
        self.master.destroy()
        self.driver.quit()

        # clean log files if exist (maybe good to keep?)
        dir = os.listdir(os.getcwd())
        for item in dir:
            if item.endswith(".log"):
                os.remove(os.path.join(os.getcwd(), item))

    # from -> http://stackoverflow.com/a/12578715
    def is_windows_64bit(self):
        if 'PROCESSOR_ARCHITEW6432' in os.environ:
            return True
        return os.environ['PROCESSOR_ARCHITECTURE'].endswith('64')

    def load_data(self):
        self.loadingBar["value"] = 25
        self.master.update_idletasks()

        # connect to webdriver - Try Google Chrome, then Firefox
        chrome = False
        try:
            chromeOptions = webdriver.ChromeOptions()
            prefs = {"profile.managed_default_content_settings.images": 2,
                     "profile.managed_default_content_settings.stylesheet": 2}
            chromeOptions.add_experimental_option("prefs", prefs)

            if platform.system() == 'Darwin':
                self.driver = webdriver.Chrome(executable_path=os.getcwd() + '/bin/chrome/chromedriverDarwin',
                                               chrome_options=chromeOptions)
            elif platform.system() == 'Windows':
                self.driver = webdriver.Chrome(executable_path=os.getcwd() + '/bin/chrome/chromedriverWindows.exe',
                                               chrome_options=chromeOptions)
            elif platform.system() == 'Linux':
                self.driver = webdriver.Chrome(executable_path=os.getcwd() + '/bin/chrome/chromedriverLinux',
                                               chrome_options=chromeOptions)
            else:
                print "Fatal error - incompatible operating system"
                exit()
            chrome = True
        except:
            print "Could not load Chrome, trying Firefox."

        if not chrome:
            try:
                # setup firefox profile, no images, no css for speed
                firefox_profile = webdriver.FirefoxProfile()
                firefox_profile.add_extension(os.getcwd() + "/bin/ext/quickjava-2.1.2-fx.xpi")
                firefox_profile.set_preference("thatoneguydotnet.QuickJava.curVersion",
                                               "2.1.2.1")  ## Prevents loading the 'thank you for installing screen'
                firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Images",
                                               2)  ## Turns images off
                firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.AnimatedImage",
                                               2)  ## Turns animated images off
                firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.CSS", 2)  ## CSS
                firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Flash", 2)  ## Flash
                firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Java", 2)  ## Java
                firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Silverlight",
                                               2)  ## Silverlight

                # load driver according to operating system,
                if platform.system() == 'Darwin':
                    self.driver = webdriver.Firefox(executable_path=os.getcwd() + '/bin/gecko/geckodriverDarwin',
                                                    firefox_profile=firefox_profile)
                elif platform.system() == 'Windows':
                    if self.is_windows_64bit():
                        self.driver = webdriver.Firefox(
                            executable_path=os.getcwd() + '/bin/gecko/geckodriverWindows64.exe',
                            firefox_profile=firefox_profile)
                    else:
                        self.driver = webdriver.Firefox(
                            executable_path=os.getcwd() + '/bin/gecko/geckodriverWindows.exe',
                            firefox_profile=firefox_profile)
                elif platform.system() == 'Linux':
                    self.driver = webdriver.Firefox(executable_path=os.getcwd() + '/bin/gecko/geckodriverLinux',
                                                    firefox_profile=firefox_profile)
                else:
                    print "Fatal error - incompatible operating system"
                    exit()
            except:
                # get no firefox exception
                print "You must have Firefox or Google Chrome installed"
                exit(6)

        # hide window/throw in corner or show
        self.browser_show()
        self.master.lift()

        self.loadingBar["value"] = 50
        self.master.update_idletasks()

        # load intial site
        self.driver.get("http://libcal.lakeheadu.ca/rooms_acc.php?gid=13445")
        assert "The Chancellor Paterson Library" in self.driver.title

        # scrape dates, only have to do once
        self.tupleDates = collections.OrderedDict()
        dateWheel = Select(self.driver.find_element_by_id("datei"))
        for date in dateWheel.options:
            self.tupleDates[date.text] = date.get_attribute("value")

        self.loadingBar["value"] = 75
        self.master.update_idletasks()

        # remove loading option, pull dates from tuple to list
        # set chosen date to top entry in date list
        # format: {'Monday June 1st': '06-01-17'}
        self.availDates.remove(self.loadingMsg)
        for key, value in self.tupleDates.iteritems():
            self.availDates.append(key)
        self.chosenDate.set(self.availDates[0])

        # get menu, clear and add dates
        menu = self.dateOptionMenu["menu"]
        menu.delete(0, "end")
        for val in self.availDates:
            menu.add_command(label=val, command=lambda pvalue=val: self.date_click(pvalue))

        # update GUI, will show current date in selection
        self.master.update()
        # initial load data
        self.date_click()
        # make sure on top
        self.master.lift()

    def date_click(self, value=""):
        # do nothing if same day clicked
        if value:
            if self.chosenDate.get() == value:
                return
            else:
                self.chosenDate.set(value)
                self.loadingBar["value"] = 20
                self.submit["state"] = "disabled"
                self.master.update_idletasks()

        # clear timeOptionList contents, showing loading message
        self.timeOptionList.delete(0, END)
        self.timeOptionList.insert(0, self.loadingMsg)
        self.timeOptionList.config(state=DISABLED)
        if value:
            self.loadingBar["value"] = 40
        self.master.update()

        # load selected date webpage
        self.driver.get(
            "http://libcal.lakeheadu.ca/rooms_acc.php?gid=13445&d=" + self.tupleDates[self.chosenDate.get()] + "&cap=0")

        # does this wait until page is loaded?
        assert "The Chancellor Paterson Library" in self.driver.title
        # make sure loaded?

        # create 2D array of [0]'s
        self.roomTimeList = [[0] for _ in range(16)]

        if value:
            self.loadingBar["value"] = 60
            self.master.update_idletasks()
        self.room_click("room")

    def room_click(self, value=""):
        if value == "room":
            self.loadingBar["value"] = 95
        else:
            self.loadingBar["value"] = 30
            self.submit["state"] = "disabled"
            self.master.update_idletasks()
        # set roomIndex to new room
        self.roomIndex = self.availRooms.index(self.chosenRoom.get())

        # clear timeOptionList contents, showing loading message
        self.timeOptionList.delete(0, END)
        self.timeOptionList.insert(0, self.loadingMsg)
        self.timeOptionList.config(state=DISABLED)
        if value != "room":
            self.loadingBar["value"] = 40
        self.master.update()

        form = self.driver.find_element_by_id("roombookingform")
        if value != "room":
            self.loadingBar["value"] = 60
            self.master.update_idletasks()
        rooms = form.find_elements_by_tag_name("fieldset")
        if value != "room":
            self.loadingBar["value"] = 80
            self.master.update_idletasks()

        # don't scrape again if already loaded
        if self.roomTimeList[self.roomIndex] == [0]:
            for room in rooms:
                # find selected room
                if room.find_element_by_tag_name("h2").text[:7] != self.chosenRoom.get():
                    continue
                # go through each room, pulling times to array
                times_out = []
                check_boxes = room.find_elements_by_class_name("checkbox")

                for box in check_boxes:
                    # output checkbox text
                    times_out.append(self.availTimes.index(box.find_element_by_tag_name("label").text))
                self.roomTimeList.insert(self.roomIndex, times_out)
                break

        # for the selected room, set time slots
        self.timeOptionList.config(state=NORMAL)
        self.timeOptionList.delete(0, END)
        if len(self.roomTimeList[self.roomIndex]) > 0:
            for timeSlot in self.roomTimeList[self.roomIndex]:
                self.timeOptionList.insert(END, self.availTimes[timeSlot])
        else:
            self.timeOptionList.insert(END, "No times available")
            self.timeOptionList.config(state=DISABLED)

        # update height, update the whole list if it isnt first run
        self.timeOptionList.configure(height=len(self.roomTimeList[self.roomIndex]))

        # colorize alternating lines of the listbox
        for i in range(0, len(self.roomTimeList[self.roomIndex]), 2):
            self.timeOptionList.itemconfigure(i, background='#f0f0ff')

        self.loadingBar["value"] = 100
        self.master.update_idletasks()

        # update GUI size and remove loading bar
        self.submit["state"] = "normal"
        self.master.update()

    def submit_click(self, value=""):
        if self.emailEntry.get().strip()[-13:] != "@lakeheadu.ca":
            tkMessageBox.showerror("Email format error", "Please make sure to use a valid @lakeheadu.ca email address")
            return

        self.save_data()

        selection = self.timeOptionList.curselection()
        if selection == ():
            return
        if len(selection) > 4 and self.overrideVal.get() != 1:
            tkMessageBox.showerror("2hr Limit", "Sorry, you can only book 2hrs per day")
            return
        if self.confirmVal.get() == 1:
            outputTimes = ""
            lineBreak = '-' * (len(self.chosenDate.get()) + 2)
            for index in selection:
                pIndex = self.roomTimeList[self.roomIndex][int(index)]
                outputTimes += self.availTimes[pIndex] + "\n"
            if not tkMessageBox.askokcancel("Please confirm the following times",
                                            self.chosenDate.get() + "\n" + lineBreak + "\n" +
                                                    self.availRooms[self.roomIndex] + "\n" + outputTimes):
                return

        # if output, get times
        self.outputTimeArray = []
        for index in selection:
            pIndex = self.roomTimeList[self.roomIndex][int(index)]
            self.outputTimeArray.append(self.availTimes[pIndex])

        # book rooms
        self.book_times()

        # check if any rooms were unavailable
        if len(self.outputTimeArray) > 0:
            outputText = "The following times were unavailable to book\n----------------------------------------\n"
            for item in self.outputTimeArray:
                outputText += item + "\n"
            tkMessageBox.showerror("Unavailable", outputText)

        # clear booked room time slots,
        self.roomTimeList[self.roomIndex] = [0]
        self.room_click()

    def book_times(self):
        # refresh page
        try:
            self.driver.refresh()
        except:
            print "Could not load the browser"
        assert "The Chancellor Paterson Library" in self.driver.title

        outputText = "Successfully booked the following rooms \n" + self.chosenDate.get() + "\n" + '-' * (
            len(self.chosenDate.get()) + 2) + "\n" + self.chosenRoom.get()
        outputLength = 0
        outputRandomId = []
        self.outputEmail = self.emailEntry.get().strip()[:-13] + "@lakeheadu.ca"
        while True:
            # get rooms
            form = self.driver.find_element_by_id("roombookingform")
            rooms = form.find_elements_by_tag_name("fieldset")

            for room in rooms:
                # find selected room
                if room.find_element_by_tag_name("h2").text[:7] != self.availRooms[self.roomIndex]:
                    continue

                consect = 0
                selectedTimes = []
                checkBoxes = room.find_elements_by_class_name("checkbox")
                for box in checkBoxes:
                    # loops through times from start every time, how can I keep place of where left off?
                    timeSlot = box.find_element_by_tag_name("label")
                    if timeSlot.text in self.outputTimeArray:
                        consect += 1
                        # adding another time will not overflow 2hrs, add it
                        if consect <= 4:
                            # check for consecutive, if array is blank, add time regardless
                            # try the simplify
                            if selectedTimes != []:
                                for selectedIndex in selectedTimes:
                                    # check if time is in consecutive order
                                    # can I put or in between two conditions????
                                    if (self.availTimes.index(timeSlot.text) - 1) == selectedIndex or (
                                                self.availTimes.index(timeSlot.text) + 1) == selectedIndex:
                                        outputText += "\n" + timeSlot.text
                                        # add to selectedTimes
                                        selectedTimes.append(self.availTimes.index(timeSlot.text))
                                        # remove timeSlot from array
                                        self.outputTimeArray.remove(timeSlot.text)
                                        # click timeSlot
                                        timeSlot.click()
                                        break
                            else:
                                outputText += "\n" + timeSlot.text
                                # add to selectedTimes
                                selectedTimes.append(self.availTimes.index(timeSlot.text))
                                # remove timeSlot from array
                                self.outputTimeArray.remove(timeSlot.text)
                                # click timeSlot
                                timeSlot.click()

                        # if 4 boxes are selected, break and book the selected rooms
                        else:
                            break
                    # if at least one box is selected, book it, must be consecutive
                    elif consect > 0:
                        break

                break
            if consect > 0:
                # fill info of form
                self.driver.find_element_by_id("fname").send_keys(self.fnameEntry.get())
                self.driver.find_element_by_id("lname").send_keys(self.lnameEntry.get())

                # always uses random_id, override 2hrs just prevents user from booking more
                # then 2hrs at a time
                randId = self.id_generator()
                outputRandomId.append(randId)
                self.driver.find_element_by_id("email").send_keys(self.emailEntry.get().strip()[:-13] + "+" + randId + "@lakeheadu.ca")

                # submit
                self.driver.find_element_by_id("s-lc-rm-ac-but").click()

                # assert it is success
                try:
                    # search for success element / fail
                    # handle failure
                    element = WebDriverWait(self.driver, 10).until(
                        EC.visibility_of_element_located((By.ID, "s-lc-rm-ac-success"))
                    )
                except TimeoutException:
                    tkMessageBox.showerror("ERROR", "Error booking the room. Exiting...")
                    self.driver.quit()
            # refresh page
            self.driver.refresh()

            # if out of times to book or cannot book times, break out
            if len(self.outputTimeArray) <= 0 or len(self.outputTimeArray) == outputLength:
                break

            outputLength = len(self.outputTimeArray)

        # print success message
        if len(self.outputTimeArray) <= 0:
            print outputRandomId
            tkMessageBox.showinfo("Success", outputText)
            # ask to confirm
            # if email used to book rooms does not match email authEmail
            ## show message that redirection is happening
            if tkMessageBox.askyesno("Confirm times", "Would you like to confirm your booked times?"):
                # check if
                if self.userInfo['authEmail'] != self.outputEmail:
                    tkMessageBox.showinfo("Redirecting...", "Press OK and login with " + self.outputEmail + " in the browser that will open." )
                self.email_click(outputRandomId)

    # random ID generator -> http://stackoverflow.com/a/2257449
    def id_generator(self, size=8, chars=string.ascii_uppercase + string.digits):
        return ''.join(random.choice(chars) for _ in range(size))


root = Tk()
libGui = GUI(root)
root.mainloop()
