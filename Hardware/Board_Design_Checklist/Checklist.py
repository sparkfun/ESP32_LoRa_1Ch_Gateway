#!/usr/bin/env python
# -*- coding: utf-8 -*-
# vim:sw=4:ts=4:et
# 12/3/2014
# SparkFun Electronics
# Pete Lewis
# Design Review Checklist Python Module

# Description ///////////////////////////////////////

# This is a little GUI (using Tkinter) intended to make reviewing a board design a bit easier and faster.

# I have included the necessary modules and a batch file to make installation easier.
# First, install Python 27 onto your machine.
# (https://www.python.org/ftp/python/2.7.9/python-2.7.9.msi)
# Note, It must be installed in the directory "c:\Python27\" for the batch file installation to work.
# Clone the repo.
# Run the batch file named "Install_Modules.bat" from within the repo.
# You will see a command prompt open up and a bunch of readout.
# When it's done, you should have the three necessary modules installed.

# You can also download each module individually and install them yourself if you like:

# https://pypi.python.org/pypi/xlrd
# https://pypi.python.org/pypi/xlwt
# https://pypi.python.org/pypi/xlutils

# Place the PCB_Checklist_v10.xls into the same folder as this module and run it.
# Note, the XLS file can be named anything, just make sure it's the only ".xls" file in the same directory.
# I alter the xls file to include the product I'm going to review.
# For example: "PCB_Checklist_v10_EL_Sequencer_v13.xls"
# It will find the first file in the same directory with an ".xls" extension and begin reading/writing to that file.
# It will prompt you all of the questions in the excell doc (one at a time), and you can click "YES", "NO", "n/a", "back" or "skip".
# It will then write to the appropriate cell in the XLS file and put your response (and username)
# You can write a comment in the comment box before you click a button,
# and it will add this comment to your response in the appropriate cell in the excell doc.

# Imports and variables ///////////////////////////////////////

# Look in current directory and find the xls file
# If there is a review "xls" file in progress (that is, it's a new file name with a "board" and "version" ammended to it),
# then jump to continue working on that file, rather than starting a fresh review.
import glob
import getpass
import os
for file in glob.glob("*.xls"):
        #print("found file: " + file)
        if(file != "Board_Design_Checklist.xls"):
            print "Additional XLS file found. Jumping to edit that now."
            excel_file_name = file
            output_file_name = excel_file_name
        else:
            excel_file_name = "Board_Design_Checklist.xls"
            output_file_name = excel_file_name

# get current User Name (should work in Windows and Linux)
current_user = getpass.getuser()
print "Current User is " + current_user
print "Current Review File is " + output_file_name

# for reading excel docs
import xlrd
book = xlrd.open_workbook(excel_file_name)
#print "The number of worksheets is", book.nsheets
#print "Worksheet name(s):", book.sheet_names()
sh = book.sheet_by_index(0)
#print sh.name, sh.nrows, sh.ncols
global current_check
current_check = "blank"
global section
section = "blank"

#for writing to excell docs
import xlwt
import time
font0 = xlwt.Font()
font0.name = 'Times New Roman'
font0.colour_index = 3 # 2 is red, 3 is green
font0.bold = True
style0 = xlwt.XFStyle()
style0.font = font0
wb = xlwt.Workbook()
user_answer = "BLANK"


from xlutils.copy import copy
from xlrd import *
w = copy(open_workbook(excel_file_name))
global first_check
first_check = 17

global check_number
check_number = first_check

global is_relevant
is_relevant = 1


# GUI setup (prompt, comment box, and buttons) ///////////////////////////////////////

from Tkinter import *

root = Tk()

is_revision = IntVar()
has_pth = IntVar()
has_i2c = IntVar()
has_SPI = IntVar()
has_1x6_ftdi_header = IntVar()
has_jumpers = IntVar()
has_leds = IntVar()
has_vreg = IntVar()
has_QFN = IntVar()
has_SS = IntVar()
has_multi_sheet = IntVar()
is_kit = IntVar()
has_serial = IntVar()
is_4layer = IntVar()
has_milling = IntVar()
has_switchorbutton = IntVar()
has_collaborator = IntVar()
board_name = StringVar()
board_ver = StringVar()

def write_dependancies():
    #print("revision: %d" % (is_revision.get()))
    #print("has PTH: %d" % (has_pth.get()))
    #print("has_i2c: %d" % (has_i2c.get()))
    #print("has_SPI: %d" % (has_SPI.get()))
    #print("has_1x6_ftdi_header: %d" % (has_1x6_ftdi_header.get()))
    #print("has_jumpers: %d" % (has_jumpers.get()))
    #print("has_leds: %d" % (has_leds.get()))
    #print("has_vreg: %d" % (has_vreg.get()))
    #print("has_QFN: %d" % (has_QFN.get()))
    #print("has_SS: %d" % (has_SS.get()))
    #print("has_multi_sheet: %d" % (has_multi_sheet.get()))
    #print("kit: %d" % (is_kit.get()))
    w.get_sheet(0).write(1,2,is_revision.get())
    w.get_sheet(0).write(2,2,has_pth.get())
    w.get_sheet(0).write(3,2,has_i2c.get())
    w.get_sheet(0).write(4,2,has_SPI.get())
    w.get_sheet(0).write(5,2,has_1x6_ftdi_header.get())
    w.get_sheet(0).write(6,2,has_jumpers.get())
    w.get_sheet(0).write(7,2,has_leds.get())
    w.get_sheet(0).write(8,2,has_vreg.get())
    w.get_sheet(0).write(9,2,has_QFN.get())
    w.get_sheet(0).write(10,2,has_SS.get())
    w.get_sheet(0).write(11,2,has_multi_sheet.get())
    w.get_sheet(0).write(12,2,is_kit.get())
    w.get_sheet(0).write(13,2,has_serial.get())
    w.get_sheet(0).write(14,2,is_4layer.get())
    w.get_sheet(0).write(15,2,has_milling.get())
    w.get_sheet(0).write(16,2,has_switchorbutton.get())
    w.get_sheet(0).write(17,2,has_collaborator.get())
    w.get_sheet(0).col(1).width = 256 * 100 # set column 2 to 100 characters wide
    global output_file_name
    output_file_name = excel_file_name[0:-4] + "_" + frame01.boardNameBox.get() + "_" + \
            frame01.boardVerBox.get() + ".xls"
    w.save(output_file_name)
    book = xlrd.open_workbook(output_file_name) # need to re-open it to get these changes we just saved. Not sure why, but it doesn't see changes unless I call this here.
    sh = book.sheet_by_index(0)

def create_summary():
    print "Creating summary..."
    summary_file_name = output_file_name[0:-4] + "_summary.txt"
    summary = open(summary_file_name, "w")
    issue_num = 1
    book = xlrd.open_workbook(output_file_name) # need to re-open it to get these changes we just saved. Not sure why, but it doesn't see changes unless I call this here.
    sh = book.sheet_by_index(0)
    for counter in range(1,sh.nrows - 1):   # Don't read past last row (0 index)
        comment = sh.cell_value(rowx=counter, colx=7)  # pull in a comment from the current row
        check = sh.cell_value(rowx=counter, colx=1)
        answer = sh.cell_value(rowx=counter, colx=6)
        if(comment != ""): # if it ain't blank, then that means there is a comment and we want to write it to the summary text file
        #print comment
            summary.write(str(issue_num) + ") " + check + " [" + answer + "]\ncomment: " + comment + "\n\n")
            issue_num = issue_num + 1
        print "complete :)"
    summary.close()

class MyFrame(Frame):
    def __init__(self):
        Frame.__init__(self)
        #self.master.geometry("1000x500")
        self.master.title("File: " + excel_file_name)
        self.grid()
 
        self.prompt = Label(self, text = "Please check all that apply...", width = 100, font=("Arial", 16), anchor=W, justify=LEFT)
        self.prompt.grid(row= 0, column = 0, columnspan=2)
 
        self.c1 = Checkbutton(self, text="Revision", variable = is_revision)
        self.c1.grid(row=1, column = 0, sticky=W)
        self.c2 = Checkbutton(self, text="PTH", variable = has_pth)
        self.c2.grid(row=2, column = 0, sticky=W)
        self.c3 = Checkbutton(self, text="I2C", variable = has_i2c)
        self.c3.grid(row=3, column = 0, sticky=W)
        self.c4 = Checkbutton(self, text="SPI", variable = has_SPI)
        self.c4.grid(row=4, column = 0, sticky=W)
        self.c5 = Checkbutton(self, text="1x6 FTDI header", variable = has_1x6_ftdi_header)
        self.c5.grid(row=5, column = 0, sticky=W)
        self.c6 = Checkbutton(self, text="Jumpers", variable = has_jumpers)
        self.c6.grid(row=6, column = 0, sticky=W)
        self.c7 = Checkbutton(self, text="LEDS", variable = has_leds)
        self.c7.grid(row=7, column = 0, sticky=W)
        self.c8 = Checkbutton(self, text="Vreg", variable = has_vreg)
        self.c8.grid(row=8, column = 0, sticky=W)
        self.c9 = Checkbutton(self, text="QFN, BGA, or small package type", variable = has_QFN)
        self.c9.grid(row=1, column = 1, sticky=W)
        self.c10 = Checkbutton(self, text="Selective Soldering", variable = has_SS)
        self.c10.grid(row=2, column = 1, sticky=W)
        self.c11 = Checkbutton(self, text="Multiple Sheets in schematic", variable = has_multi_sheet)
        self.c11.grid(row=3, column = 1, sticky=W)
        self.c12 = Checkbutton(self, text="KIT", variable = is_kit)
        self.c12.grid(row=4, column = 1, sticky=W)
        self.c13 = Checkbutton(self, text="Serial Com", variable = has_serial)
        self.c13.grid(row=5, column = 1, sticky=W)
        self.c14 = Checkbutton(self, text="4-layer", variable = is_4layer)
        self.c14.grid(row=6, column = 1, sticky=W)
        self.c15 = Checkbutton(self, text="Milling", variable = has_milling)
        self.c15.grid(row=7, column = 1, sticky=W)
        self.c16 = Checkbutton(self, text="Switches or Buttons", variable = has_switchorbutton)
        self.c16.grid(row=8, column = 1, sticky=W)
        self.c17 = Checkbutton(self, text="Collaborator **NEW**", font=("Arial", 10, "bold", "italic"), variable = has_collaborator)
        self.c17.grid(row=9, column = 0, sticky=W)        
 
        self.boardNameFrame = LabelFrame(self, bd = 0)
        self.boardNameFrame.grid(row = 10, column = 0, sticky = W)
        self.boardNameBoxLabel = Label(self.boardNameFrame, text = "Board Name:")
        self.boardNameBoxLabel.pack(side = LEFT)
        self.boardNameBox = Entry(self.boardNameFrame, textvariable =
                                  board_name, width = 60)
        self.boardNameBox.pack(side = LEFT)
 
 
        self.boardVerFrame = LabelFrame(self, bd = 0)
        self.boardVerFrame.grid(row = 10, column = 1, sticky = W)
        self.boardVerBoxLabel = Label(self.boardVerFrame, text = "Board Version:")
        self.boardVerBoxLabel.pack(side = LEFT)
        self.boardVerBox = Entry(self.boardVerFrame, textvariable = board_ver)
        self.boardVerBox.pack(side = LEFT)
 
        self.submit1 = Button(self, text='Submit', command = self.submit_dependancies)
        self.submit1.grid(row=11, sticky=W, pady=4)

        #Check to see if we are working with a previous review name already...
        global output_file_name
        if(output_file_name != "Board_Design_Checklist.xls"):
            self.submit_continue_previous_review()
            
        return
 
    def set_page_1(self):
        # destroy all checkboxes and button from page 0
        self.prompt.destroy()
        self.c1.destroy()
        self.c2.destroy()
        self.c3.destroy()
        self.c4.destroy()
        self.c5.destroy()
        self.c6.destroy()
        self.c7.destroy()
        self.c8.destroy()
        self.c9.destroy()
        self.c10.destroy()
        self.c11.destroy()
        self.c12.destroy()
        self.c13.destroy()
        self.c14.destroy()
        self.c15.destroy()
        self.c16.destroy()
        self.c17.destroy()
        self.submit1.destroy()
        self.boardVerBox.destroy()
        self.boardVerBoxLabel.destroy()
        self.boardNameBox.destroy()
        self.boardNameBoxLabel.destroy()
        self.boardNameFrame.destroy()
 
        # add in section, prompt, comment text/box, and buttons to click in answers
 
        self.section = Label(self, text = "Section: " + section, width = 100, font=("Arial", 12), anchor=W, justify=LEFT)
        self.section.grid(row= 0, column = 0, columnspan=3)
 
        self.prompt = Label(self, text = current_check, width = 100, font=("Arial", 16), anchor=W, justify=LEFT)
        self.prompt.grid(row= 1, column = 0, columnspan=3)
 
        self.comment_text = Label(self, text = "comment (optional):", width = 100, font=("Arial", 12), anchor=W, justify=LEFT)
        self.comment_text.grid(row= 2, column = 0, columnspan=2)
 
        self.input = Entry(self, width = 150)
        self.input.grid(row= 3, column = 0, columnspan=3, sticky=W)
 
        self.b1 = Button(self, text = "YES", command = self.submit_click_yes)
        self.b1.grid(row= 4, column = 0, sticky=W)
 
        self.b2 = Button(self, text = "NO", command = self.submit_click_no)
        self.b2.grid(row= 5, column = 0, sticky=W)
 
        self.b3 = Button(self, text = "n/a", command = self.submit_click_na)
        self.b3.grid(row= 6, column = 0, sticky=W)
 
        self.b4 = Button(self, text = "go back", command = self.submit_click_back)
        self.b4.grid(row= 7, column = 0, sticky=W)
 
        self.b5 = Button(self, text = "skip", command = self.submit_click_skip)
        self.b5.grid(row= 8, column = 0, sticky=W)
 
        self.b6 = Button(self, text = "create summary", command = self.submit_click_summary)
        self.b6.grid(row= 8, column = 3, sticky=W)
        return
 
    def update_frame(self):
        self.section = Label(self, text = "[" + section + "]", width = 100, font=("Arial", 16), anchor=W, justify=LEFT)
        self.section.grid(row= 0, column = 0, columnspan=3)
        self.prompt = Label(self, text = current_check, width = 100, font=("Arial", 16), anchor=W, justify=LEFT)
        self.prompt.grid(row= 1, column = 0, columnspan=3)
        return
 
 
    def update_doc(self):
        w.get_sheet(0).write(check_number, 5, current_user)
        w.get_sheet(0).write(check_number, 6, user_answer)
        w.get_sheet(0).write(check_number, 7, self.input.get())
        w.get_sheet(0).col(1).width = 256 * 100 # set column 2 to 100 characters wide
        w.save(output_file_name)
        return
 
    def check_if_relevant(self):
        global check_number
        global is_relevant
        global user_answer
        book = xlrd.open_workbook(output_file_name) # need to re-open it to get these changes we just saved. Not sure why, but it doesn't see changes unless I call this here.
        sh = book.sheet_by_index(0)
        dependancy = sh.cell_value(rowx=check_number, colx=3)
        if(dependancy == ""):
            #print "No dependancy"
            is_relevant = 1
            #print "relevant"
        else:
            dependancy = int(dependancy) - 1
            #print "Dependant on macro question " + str(dependancy)
            is_relevant = sh.cell_value(rowx=dependancy, colx=2)
            is_relevant = int(is_relevant)
            #print is_relevant
            if(is_relevant):
                #print "relevant"
                do_nothing = 1 # do nothing, this is just here as a placeholder for the if statement above        
            else:
                print sh.cell_value(rowx=check_number, colx=1) + " is not relevant"
                #print "is not relevant"
                user_answer = "*n/a*"
                self.update_doc()
                return 
 
 
    def increment_prompt(self):
        global check_number
        global section
        global current_check
        global is_relevant  
        self.input.delete(0,END)
        if check_number + 2 <= sh.nrows:    # incremented val > last index ?
            check_number = check_number + 1
        self.check_if_relevant()
        #Skip irrelevant checks
        while(is_relevant == 0):
            if check_number + 2 <= sh.nrows:    # incremented val > last index ?
                check_number = check_number + 1
            self.check_if_relevant()
 
        
        current_check = sh.cell_value(rowx=check_number, colx=1)
        section = sh.cell_value(rowx=check_number, colx=0)
        self.update_frame()     
        return

    def find_where_we_left_off(self):
        print "finding where we left off..."
        global check_number
        check_number = check_number + 1
        user_input = sh.cell_value(rowx=check_number, colx=6) # colx = 6 is the 6th columb - where the user answers are stored.
        #print user_input
        while(user_input != ''):
            user_input = sh.cell_value(rowx=check_number, colx=6) # colx = 6 is the 6th columb - where the user answers are stored.
            #print user_input
            if(user_input == ''):
                    #print "BLANK"
                    check_number = check_number - 1 # This means we've found where we left off, and we need to step back one spot, so it works with increment_prompt()
                    
            else:
                check_number = check_number + 1
            
        
    def submit_click_yes(self):
        global user_answer
        user_answer = "YES"
        self.update_doc()
        self.increment_prompt()
        return
 
    def submit_click_no(self):
        global user_answer
        user_answer = "NO"
        self.update_doc()
        self.increment_prompt()
        return
 
    def submit_click_na(self):
        global user_answer
        user_answer = "n/a"
        self.update_doc()
        self.increment_prompt()
        return
 
    def submit_click_back(self):
        global check_number
        global is_relavant
        global first_check
        if(check_number != (first_check + 1)): # to avoid going back beyond the first check - into the macro variables
            check_number = check_number - 1 # go back one step
            self.check_if_relevant() # see if it's relevant
            #Skip irrelevant checks
            while(is_relevant == 0):
                check_number = check_number - 1
                self.check_if_relevant()
 
            check_number = check_number - 1 # go back another, then call increment (which will add one back)
            self.increment_prompt()
        return
 
    def submit_click_skip(self):
        global check_number
        if check_number + 2 <= sh.nrows:    # incremented val > last index ?
            check_number = check_number + 1
        self.increment_prompt()
        return
 
    def submit_dependancies(self):
        write_dependancies()
        self.set_page_1()
        self.increment_prompt()
        return
    
    def submit_continue_previous_review(self):
        self.set_page_1()
        self.find_where_we_left_off()
        self.increment_prompt()
        return
 
    def submit_click_summary(self):
        create_summary()
        return


# Actually run the GUI ///////////////////////////////////////

frame01 = MyFrame()
frame01.mainloop()
