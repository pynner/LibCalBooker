from Tkinter import *
import tkMessageBox
import platform
import json
import string
import random
import collections
import os
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile


class GUI:
    def __init__(self, master):
        self.master = master
        self.master.title("LibCal Booker v1.0.0")

        print platform.system()

        ################-GLOBAL-VARS-############################
        self.online = True
        self.phantom = False
        self.driver = ""
        self.tupleDates = []
        self.roomTimeList = []
        self.roomIndex = 0
        self.outputTimeArray = []
        self.userInfo = []

        ###############-LOAD-FILE-###############################
        try:
            with open("userInfo.json") as data_file:
                self.userInfo = json.load(data_file)
        except:
            print "No existing user"
            with open("userInfo.json", "w+") as data_file:
                self.userInfo = dict(first="Mike", last="Anderson", email="mand@lakeheadu.ca")
                json.dump(self.userInfo, data_file)

        ###################-DATE SELECTION-######################
        # set up availDates
        self.loadingMsg = "Loading......................."
        self.availDates = [self.loadingMsg]
        # set up chosen date and give default value
        self.chosenDate = StringVar(self.master)
        self.chosenDate.set(self.loadingMsg)
        # set up dateOptionMenu
        self.dateOptionMenu = OptionMenu(self.master, self.chosenDate, *self.availDates, command=self.date_click)
        self.dateOptionMenu.grid(row=0, column=0, columnspan=1, sticky=(N, S, E, W), pady=(8,0))

        ##################-ROOM SELECTION-######################
        # set up availRooms
        self.availRooms = ['LI 1004', 'LI 1006', 'LI 1007', 'LI 1008', 'LI 1009', 'LI 1010', 'LI 4001', 'LI 4002',
                           'LI 4003', 'LI 4004', 'LI 4005', 'LI 4006', 'LI 4007', 'LI 4008', 'LI 4009', 'LI 4010']
        self.availTimes = ['8:00am - 8:30am', '8:30am - 9:00am', '9:00am - 9:30am', '9:30am - 10:00am', '10:00am - 10:30am',
                           '10:30am - 11:00am', '11:00am - 11:30am', '11:30am - 12:00pm', '12:00pm - 12:30pm', '12:30pm - 1:00pm',
                           '1:00pm - 1:30pm', '1:30pm - 2:00pm', '2:00pm - 2:30pm', '2:30pm - 3:00pm', '3:00pm - 3:30pm',
                           '3:30pm - 4:00pm', '4:00pm - 4:30pm', '4:30pm - 5:00pm', '5:00pm - 5:30pm', '5:30pm - 6:00pm', '6:00pm - 6:30pm',
                           '6:30pm - 7:00pm', '7:00pm - 7:30pm', '7:30pm - 8:00pm', '8:00pm - 8:30pm', '8:30pm - 9:00pm', '9:00pm - 9:30pm',
                           '9:30pm - 10:00pm', '10:00pm - 10:30pm', '10:30pm - 11:00pm', '11:00pm - 11:30pm', '11:30pm - 11:59pm']
        # set up chosen room and give default value
        self.chosenRoom = StringVar(self.master)
        self.chosenRoom.set(self.availRooms[0])
        # set up roomOptionMenu
        self.roomOptionMenu = OptionMenu(master, self.chosenRoom, *self.availRooms, command=self.room_click)
        self.roomOptionMenu.grid(row=1, column=0, columnspan=1, sticky=(N, S, E, W))

        ##################-TIME SELECTION-####################
        self.timeOptionList = Listbox(self.master, selectmode=MULTIPLE, height=20, exportselection=0)
        self.timeOptionList.grid(row=2, column=0, rowspan=10, columnspan=1, sticky=(N, S, E, W), padx=5, pady=5)
        self.timeOptionList.insert(0, self.loadingMsg)

        #################-BUTTONS-##########################
        # SIDE INFO
        self.infoLabel = Label(master, text="[ U S E R  I N F O ]", font=("Helvetica", 16, "bold"))
        self.infoLabel.grid(row=0, column=1, columnspan=2, sticky=E + W)

        # First Name label and input
        self.fnameLabel = Label(master, text="First: ", font=("Helvetica", 12, "bold"))
        self.fnameLabel.grid(row=1, column=1, sticky=W)

        self.fnameEntry = Entry(master)
        self.fnameEntry.grid(row=1, column=2, stick=E + W, padx=(0, 5))
        self.fnameEntry.insert(0, self.userInfo["first"])

        # Last name label and input
        self.lnameLabel = Label(master, text="Last: ", font=("Helvetica", 12, "bold"))
        self.lnameLabel.grid(row=2, column=1, sticky=W)

        self.lnameEntry = Entry(master)
        self.lnameEntry.grid(row=2, column=2, stick=E + W, padx=(0, 5))
        self.lnameEntry.insert(0, self.userInfo["last"])

        # Email label and entry
        self.emailLabel = Label(master, text="Email: ", font=("Helvetica", 12, "bold"))
        self.emailLabel.grid(row=3, column=1, sticky=W)

        self.emailEntry = Entry(master,)
        self.emailEntry.grid(row=3, column=2, stick=E + W, padx=(0, 5))
        self.emailEntry.insert(0, self.userInfo["email"])

        # Override checkbox
        self.overrideVal = IntVar(master)
        self.override = Checkbutton(master, text="Override 2hr max", variable=self.overrideVal,
                                    onvalue=1, offvalue=0, font=("Helvetica", 12))
        self.override.grid(row=4, column=2, sticky=W)

        self.confirmVal = IntVar(master)
        self.confirmVal.set(1)
        self.confirm = Checkbutton(master, text="Enable confirm dialog", variable=self.confirmVal,
                                   onvalue=1, offvalue=0, font=("Helvetica", 12))
        self.confirm.grid(row=5, column=2, sticky=W)

        # submit button
        self.submit = Button(master, text="Submit", command=self.submit_click)
        self.submit.grid(row=6, column=2, sticky=(N, S, E, W), padx=(0, 5))

        # update skeleton GUI, then load data
        self.master.update()
        self.load_data()

        # make sure window on top
        self.master.lift()

    def load_data(self):
        # connect to webdriver - PhantomJS for headless browser
        # FIREFOX faster
        if not self.phantom:
            # setup firefox profile, no images, no css for speed
            firefox_profile = webdriver.FirefoxProfile()
            firefox_profile.add_extension(os.getcwd() + "/ext/quickjava-2.1.2-fx.xpi")
            firefox_profile.set_preference("thatoneguydotnet.QuickJava.curVersion", "2.1.2.1")  ## Prevents loading the 'thank you for installing screen'
            firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Images", 2)  ## Turns images off
            firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.AnimatedImage", 2)  ## Turns animated images off
            firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.CSS", 2)  ## CSS
            firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Flash", 2)  ## Flash
            firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Java", 2)  ## Java
            firefox_profile.set_preference("thatoneguydotnet.QuickJava.startupStatus.Silverlight", 2)  ## Silverlight

            if platform.system() == 'Darwin':
                self.driver = webdriver.Firefox(executable_path=os.getcwd() + '/gecko/geckodriverOSX', firefox_profile=firefox_profile)
            elif platform.system() == 'Windows':
                self.driver = webdriver.Firefox(executable_path=os.getcwd() + '/gecko/geckodriverWINDOWS.exe', firefox_profile=firefox_profile)
            else:
                print "Fatal error - In compatible operating system"
                exit()

            # hide window/throw in corner
            self.driver.set_window_position(3000,3000)
            self.driver.set_window_size(100, 300)
        else:
            self.driver = webdriver.PhantomJS()

        if not self.online:
            self.driver.get("file:///Users/mitchellpynn/Desktop/LIB/MobileLIB_2.html")
        else:
            self.driver.get("http://libcal.lakeheadu.ca/rooms_acc.php?gid=13445")
        assert "The Chancellor Paterson Library" in self.driver.title

        # put window on top of webdriver, don't need for PhantomJS
        if not self.phantom:
            self.master.lift()

        # scrape dates, only have to do once
        self.tupleDates = collections.OrderedDict()
        dateWheel = Select(self.driver.find_element_by_id("datei"))
        for date in dateWheel.options:
            self.tupleDates[date.text] = date.get_attribute("value")

        # Remove loading option, pull dates from tuple to list
        # Set chosen date to top entry in date list
        # {'Monday June 1st': '06-01-17'}
        self.availDates.remove(self.loadingMsg)
        for key, value in self.tupleDates.iteritems():
            self.availDates.append(key)
        self.chosenDate.set(self.availDates[0])

        # Get menu, clear and add dates
        menu = self.dateOptionMenu["menu"]
        menu.delete(0, "end")
        for val in self.availDates:
            menu.add_command(label=val, command=lambda pvalue=val: self.date_click(pvalue))

        self.master.update()
        # Initial load times
        self.date_click()
        # Make sure on top
        self.master.lift()

    def date_click(self, value=""):
        # >Set chosenDate
        # >Load new page
        # >Pull new times from page
        # >> Only load first times, too slow otherwise
        # >Call func to update times for selected room

        # do nothing if same day clicked
        if value:
            if self.chosenDate.get() == value:
                return
            else:
                self.chosenDate.set(value)
        print self.chosenDate.get()

        # clear timeOptionList contents, showing loading message
        self.timeOptionList.delete(0, END)
        self.timeOptionList.insert(0, self.loadingMsg)
        self.master.update()

        # load selected date webpage
        if self.online:
            self.driver.get("http://libcal.lakeheadu.ca/rooms_acc.php?gid=13445&d=" + self.tupleDates[self.chosenDate.get()] + "&cap=0")

        # does this wait until page is loaded?
        assert "The Chancellor Paterson Library" in self.driver.title
        # make sure loaded?
        # take screenshot of day page - remove later
        if self.phantom:
            self.driver.save_screenshot('checkPage.png')

        # create 2D array of [0]'s
        self.roomTimeList = [[0] for _ in range(16)]

        self.room_click()

    def room_click(self, value=""):
        # set roomIndex to new room
        self.roomIndex = self.availRooms.index(self.chosenRoom.get())

        # clear timeOptionList contents, showing loading message
        self.timeOptionList.delete(0, END)
        self.timeOptionList.insert(0, self.loadingMsg)
        self.master.update()

        form = self.driver.find_element_by_id("roombookingform")
        rooms = form.find_elements_by_tag_name("fieldset")

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
        self.timeOptionList.delete(0, END)
        for timeSlot in self.roomTimeList[self.roomIndex]:
            self.timeOptionList.insert(END, self.availTimes[timeSlot])

        # update height, update the whole list if it isnt first run
        self.timeOptionList.configure(height=len(self.roomTimeList[self.roomIndex]))

        # Colorize alternating lines of the listbox
        for i in range(0, len(self.roomTimeList[self.roomIndex]), 2):
            self.timeOptionList.itemconfigure(i, background='#f0f0ff')

        # Update GUI size
        self.master.update()

    def submit_click(self):
        if self.emailEntry.get().strip()[-13:] != "@lakeheadu.ca":
            tkMessageBox.showerror("Email format error", "Please make sure to use a valid @lakeheadu.ca email address")
            return

        # update userFile when submitted
        with open("userInfo.json", 'r+') as data_file:
            self.userInfo["first"] = self.fnameEntry.get()
            self.userInfo["last"] = self.lnameEntry.get()
            self.userInfo["email"] = self.emailEntry.get()

            data_file.seek(0)
            json.dump(self.userInfo, data_file)
            data_file.truncate()


        selection = self.timeOptionList.curselection()
        if selection == ():
            print "Nothing to book"
            return
        if len(selection) > 4 and self.overrideVal.get() != 1:
            tkMessageBox.showerror("2hr Limit", "Sorry, you can only book 2hrs per day")
            return
        if self.confirmVal.get() == 1:
            outputTimes = ""
            print selection
            lineBreak = '-' * (len(self.chosenDate.get()) + 2)
            for index in selection:
                pIndex = self.roomTimeList[self.roomIndex][int(index)]
                outputTimes += "\t" + self.availTimes[pIndex] + "\n"
            if not tkMessageBox.askokcancel("=-= Please confirm these times =-=", self.chosenDate.get() + "\n" + lineBreak + "\n\t" +
                    self.availRooms[self.roomIndex] + "\n" + outputTimes):
                return

        # if output, get times
        self.outputTimeArray = []
        for index in selection:
            pIndex = self.roomTimeList[self.roomIndex][int(index)]
            self.outputTimeArray.append(self.availTimes[pIndex])
        self.book_times()

        if len(self.outputTimeArray) != 0:
            outputText = "The following times were unavailable to book\n----------------------------------------\n"
            for item in self.outputTimeArray:
                outputText += item + "\n"
            tkMessageBox.showerror("Error booking rooms", outputText)

        # clear booked room time slots,
        self.roomTimeList[self.roomIndex] = [0]
        self.room_click()


    def book_times(self):
        # >Go through selected items
        # >Book in largest chunk as possible
        # >Use ID_gen on every email
        # >Check response on each booking, make sure confirmed
        # >>If not confirmed, see if maxed out hours/room already booked

        print "Booking times now. Please wait...."
        # make sure connection is alive ######## not sure if best wei
        # refresh page
        self.driver.refresh()
        assert "The Chancellor Paterson Library" in self.driver.title
        print self.outputTimeArray

        outputLength = 0
        while True:
            # get rooms
            form = self.driver.find_element_by_id("roombookingform")
            rooms = form.find_elements_by_tag_name("fieldset")

            for room in rooms:
                # find selected room
                if room.find_element_by_tag_name("h2").text[:7] != self.availRooms[self.roomIndex]:
                    continue
                print "Found room " + room.find_element_by_tag_name("h2").text[:7] + ", now booking..."

                consect = 0
                selectedTimes = []
                checkBoxes = room.find_elements_by_class_name("checkbox")
                for box in checkBoxes:
                    # loops through times from start every  time, how can I keep place of where left off?
                    timeSlot = box.find_element_by_tag_name("label")
                    if timeSlot.text in self.outputTimeArray:
                        print timeSlot.text
                        consect += 1
                        # adding another time will not overflow 2hrs, add it
                        if consect <= 4:
                            # check for consecutive, if array is blank, add time regardless
                            # try the simplify
                            if selectedTimes != []:
                                for selectedIndex in selectedTimes:
                                    # check if time is in consecutive order
                                    # can I put or in between two conditions????
                                    if (self.availTimes.index(timeSlot.text) - 1) == selectedIndex or (self.availTimes.index(timeSlot.text) + 1) == selectedIndex:
                                        # add to selectedTimes
                                        selectedTimes.append(self.availTimes.index(timeSlot.text))
                                        # remove timeSlot from array
                                        self.outputTimeArray.remove(timeSlot.text)
                                        # click timeSlot
                                        timeSlot.click()
                                        break
                            else:
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
                self.driver.find_element_by_id("email").send_keys(self.emailEntry.get().strip()[:-13] + "+" + self.id_generator() + "@lakeheadu.ca")
                print self.emailEntry.get().strip()[:-13] + "+" + self.id_generator() + "@lakeheadu.ca"

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

            print self.outputTimeArray
            print len(self.outputTimeArray)
            print outputLength
            # if out of times to book or cannot book times, break out
            if len(self.outputTimeArray) <= 0 or len(self.outputTimeArray) == outputLength:
                break

            outputLength = len(self.outputTimeArray)

    # Random ID generator -> http://stackoverflow.com/a/2257449
    def id_generator(self, size=8, chars=string.ascii_uppercase + string.digits):
        return ''.join(random.choice(chars) for _ in range(size))


root = Tk()
libGui = GUI(root)
root.mainloop()
