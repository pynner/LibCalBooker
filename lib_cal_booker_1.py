from Tkinter import *
import tkMessageBox
import time
import string
import random
import collections
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


class GUI:
    def __init__(self, master):
        self.master = master
        self.master.title("LibCal Booker v1.0.0")

        ################-GLOBAL-VARS-############################
        self.tupleDates = []
        self.roomTimeList = []
        self.roomIndex = 0
        self.outputTimeArray = []

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

        # Last name label and input
        self.lnameLabel = Label(master, text="Last: ", font=("Helvetica", 12, "bold"))
        self.lnameLabel.grid(row=2, column=1, sticky=W)

        self.lnameEntry = Entry(master)
        self.lnameEntry.grid(row=2, column=2, stick=E + W, padx=(0, 5))

        # Email label and entry
        self.emailLabel = Label(master, text="Email: ", font=("Helvetica", 12, "bold"))
        self.emailLabel.grid(row=3, column=1, sticky=W)

        self.emailEntry = Entry(master)
        self.emailEntry.grid(row=3, column=2, stick=E + W, padx=(0, 5))

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

        # load saved user data
        # right here or maybe after data load?

        # update skeleton GUI, then load data
        self.master.update()
        self.load_data()

        # make sure window on top
        self.master.lift()

    def load_data(self):
        # connect to webdriver PhantomJS for headless browser
        #self.driver = webdriver.Firefox(executable_path='/Library/Python/geckodriver')
        self.driver = webdriver.PhantomJS()
        # Move window to side / off screen
        #self.driver.set_window_position(1000, 0)
        self.driver.get("http://libcal.lakeheadu.ca/rooms_acc.php?gid=13445")

        # OFFLINE USE
        # self.driver.get("file:///Users/mitchellpynn/Desktop/LIB/MobileLIB_2.html")
        assert "The Chancellor Paterson Library" in self.driver.title


        # put window on top of webdriver, don't need after PhantomJS
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

        # Initial load times
        self.date_click()

    def date_click(self, value=""):
        # >Set chosenDate
        # >Load new page
        # >Pull new times from page
        # >Call func to update times for selected room
        if value:
            if self.chosenDate.get() == value:
                return
            else:
                self.chosenDate.set(value)

        print self.chosenDate.get()
        # load selected date webpage TRY ONLINE
        self.driver.get("http://libcal.lakeheadu.ca/rooms_acc.php?gid=13445&d=" + self.tupleDates[self.chosenDate.get()] + "&cap=0")
        assert "The Chancellor Paterson Library" in self.driver.title
        # make sure loaded?
        self.driver.save_screenshot('checkPage.png')

        # FRIKKEN SLOW
        time1 = time.time()
        self.roomTimeList = []
        form = self.driver.find_element_by_id("roombookingform")
        rooms = form.find_elements_by_tag_name("fieldset")
        for room in rooms:
            # go through each room, pulling times to array
            times_out = []
            check_boxes = room.find_elements_by_class_name("checkbox")
            for box in check_boxes:
                times_out.append(box.find_element_by_tag_name("label").text)
            self.roomTimeList.append(times_out)
        print self.roomTimeList
        time2 = time.time()
        print 'getting the times function took %0.3f ms' % ((time2 - time1) * 1000.0) #20s avg
        self.master.update()
        self.room_click()

    def room_click(self, value=""):
        if value:
            self.chosenRoom.set(value)
        # clear timeOptionList contents
        self.timeOptionList.delete(0, END)  # need to check if null??
        # get index of chosenRoom
        self.roomIndex = self.availRooms.index(self.chosenRoom.get())
        # for the selected room, get all time slots
        for time in self.roomTimeList[self.roomIndex]:
            self.timeOptionList.insert(END, time)
        # update height, update the whole list if it isnt first run
        self.timeOptionList.configure(height=len(self.roomTimeList[self.roomIndex]))
        if value:
            self.timeOptionList.update_idletasks()
        # Colorize alternating lines of the listbox
        for i in range(0, len(self.roomTimeList[self.roomIndex]), 2):
            self.timeOptionList.itemconfigure(i, background='#f0f0ff')

    def submit_click(self):
        selection = self.timeOptionList.curselection()
        if selection == ():
            print "Nothing to book"
            return
        if len(selection) > 4 and self.overrideVal.get() != 1:
            tkMessageBox.showerror("2hr Limit", "Sorry, you can only book 2hrs per day")
            return
        if self.confirmVal.get() == 1:
            outputTimes = ""
            for index in selection:
                outputTimes += "\t" + self.roomTimeList[self.roomIndex][index] + "\n"
            if not tkMessageBox.askokcancel("Confirm dialog", "=-= Please confirm these times =-=\n\t"
                    + self.availRooms[self.roomIndex] + "\n" + outputTimes):
                return

        # if output, get times
        self.outputTimeArray = []
        for index in selection:
            self.outputTimeArray.append(self.roomTimeList[self.roomIndex][index])
        print self.outputTimeArray
        self.book_times()
        self.date_click()


    def book_times(self):
        # >Go through selected items
        # >Book in largest chunk as possible
        # >Use ID_gen on every email
        # >Check response on each booking, make sure confirmed
        # >>If not confirmed, see if maxed out hours/room already booked

        print "Booking times now. Please wait...."
        # make sure connection is alive ######## not sure if best wei
        assert "The Chancellor Paterson Library" in self.driver.title

        while True:
            #        # wait until page loaded, not sure if necessary
            #       try:
            #          element = WebDriverWait(driver, 3).until(
            #             EC.presence_of_element_located((By.ID, "lname"))
            #        )
            #   except TimeoutException:
            #      print "Initial time out"
            #     driver.quit()

            # get rooms // get passed in?
            form = self.driver.find_element_by_id("roombookingform")
            rooms = form.find_elements_by_tag_name("fieldset")
            for room in rooms:
                # find selected room
                if room.find_element_by_tag_name("h2").text[:7] != self.availRooms[self.roomIndex]:
                    continue
                print "Found room, now booking."
                print room.find_element_by_tag_name("h2").text[:7]

                consect = 0
                checkBoxes = room.find_elements_by_class_name("checkbox")
                for box in checkBoxes:
                    timeSlot = box.find_element_by_tag_name("label").text
                    print timeSlot
                    if timeSlot in self.outputTimeArray:
                        consect += 1
                        if consect <= 4:
                            self.outputTimeArray.remove(timeSlot)
                            box.find_element_by_tag_name("label").click()
                            continue
                        else:
                            break
                    elif consect > 0:
                        break
                break
            if consect > 0:
                # fill info of form
                ActionChains(self.driver).send_keys(Keys.END).perform()

                self.driver.find_element_by_id("fname").send_keys(self.fnameEntry.get())
                self.driver.find_element_by_id("lname").send_keys(self.lnameEntry.get())
                self.driver.find_element_by_id("email").send_keys(
                    self.emailEntry.get()[:-13] + self.id_generator() + "@lakeheadu.ca")
                print self.emailEntry.get()[:-13] + "+" + self.id_generator() + "@lakeheadu.ca"
                # submit
                self.driver.find_element_by_id("s-lc-rm-ac-but").click()

                # assert it is success then write to database
                try:
                    # search for success element
                    element = WebDriverWait(self.driver, 5).until(
                        EC.visibility_of_element_located((By.ID, "s-lc-rm-ac-success"))
                    )
                except TimeoutException:
                    tkMessageBox.showerror("ERROR", "Error booking the room. Exiting...")
                    time.sleep(5)
                    self.driver.quit()
            # refresh page
            self.driver.refresh()

            if len(self.outputTimeArray) <= 0:
                break

    # Random ID generator -> http://stackoverflow.com/a/2257449
    def id_generator(self, size=8, chars=string.ascii_uppercase + string.digits):
        return ''.join(random.choice(chars) for _ in range(size))


root = Tk()
libGui = GUI(root)
root.mainloop()
