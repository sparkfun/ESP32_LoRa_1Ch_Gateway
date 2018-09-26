Design Review Checklist Python Module
=====================================

Original Author: Pete Lewis

Creation Date: 4/9/2015

Updated: 5/10/16

Description:
-------------
This is a little GUI (using Tkinter) intended to make reviewing a board design a bit easier and faster.
I have included the necessary modules and a batch file to make installation easier.
A bunch of helpful images are located [here](https://github.com/sparkfun/Engineering_Design_Rules_Checklist/tree/master/Board_Design_Checklist/doc_images)

Python GUI Installation:
---------------------
### Windows:

1. Install [Python 2.7.x](https://www.python.org/downloads/windows/) on your machine. At the top of this page is a link to the latest release. From there download [Windows x86 MSI installer](https://www.python.org/ftp/python/2.7.11/python-2.7.11.msi).

	1a. Note, It must be installed in the directory "c:\Python27\" for the batch file installation to work.
	
	1b. import _tkinter will fail with: "ImportError: DLL load failed: %1 is not a valid Win32 application." if you try to use the 64-bit version of Python.
2. The installer will give you the option to setup your environment variables. May as well select this. At the time of this writing, it doesn't work right, but it's still helpful.

	2a. Right click on the start menu and select the 'Control Panel'.

	2b. Searching for env will show you 'Edit the system environment variables', choose this.

	2c. Click 'Environment Variables...' at the bottom.

	2d. Double click on 'Path'.

	2e. You should see 'C:\Python27\' & 'C:\Python27\Scripts'. Append a trailing '&#92;' to Scripts to fix it.

2. Clone the [repo](https://github.com/sparkfun/Engineering_Design_Rules_Checklist).
3. Run the batch file named "Install_Modules.bat" from within the repo (Excel_Modules_Installation).
4. You will see a command prompt open up and a bunch of readout.
5. When it's done, you should have the three necessary modules installed.

### OS X:

1. Install Homebrew
	```ruby -e "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install)"```
2. Install Python
	```brew install python```
3. Install pip
	```sudo easy_install pip```
4. Install Excel Modules
	```
	sudo pip install xlrd
	
	sudo pip install xlwt
	
	sudo pip install xlutils
	```
5. Clone the [repo](https://github.com/sparkfun/Engineering_Design_Rules_Checklist).
	
### Debian/Ubuntu: 

1. Install Python
	```
	sudo apt-get install python
	```
2. Install Excel Modules
	```
	sudo apt-get install python-xlrd
	
	sudo apt-get install python-xlwt
	
	sudo apt-get install python-pip
	
	sudo pip install xlutils
	```
	
### Generic Install Directions:

1. Install [Python 2.7.x](https://www.python.org/ftp/python/2.7.9/python-2.7.9.msi) on your machine.
2. Download and install the following modules.
  * [XLRD](https://pypi.python.org/pypi/xlrd)
  * [XLWT](https://pypi.python.org/pypi/xlwt)
  * [XLUTILS](https://pypi.python.org/pypi/xlutils)

You can do this with pip, for instance, run the command `pip install xlrd` to install XLRD

Running the Checklist:
----------------------
Run Checklist.py file from a directory that contains the original file `Board_Design_Checklist.xls`, or from a directory that contains an already started product specific `.xls` file.

The python file has two modes of operation:
1. If no `.xls` file exist, the program exits.
2. If an `.xls` file exists and is named `Board_Design_Checklist.xls`, the program will generate a new `.xls` file with the product name appended, such as `Board_Design_Checklist_fancynewproduct_.xls`.  The original spreadsheet is left intact.
3. If there is already a spreadsheet with product name appended, the program will open it for continuation or review.

You can double-click Checklist.py, or run from the command line as `>Checklist.py`

For example, if `folder` contains `Board_Design_Checklist.xls`: 
```
c:\>cd folder

c:\folder>dir
 Directory of c:\folder

06/14/2017  09:31 AM            25,600 Board_Design_Checklist.xls
06/14/2017  09:31 AM            18,827 Checklist.py

c:\folder>Checklist.py
Current User is marshall.taylor
Current Review File is Board_Design_Checklist.xls

c:\folder>dir
 Directory of c:\temp

06/14/2017  09:31 AM            25,600 Board_Design_Checklist.xls
06/15/2017  04:06 PM            22,016 Board_Design_Checklist_fancynewproduct_.xls
06/14/2017  09:31 AM            18,827 Checklist.py

c:\folder>
```



The first screen is a combination of checks. Check all of the boxes that apply to your design. This will help the GUI avoid asking you unnecessary checks.  The name entered here will be appended to the spreadsheet.

The next screens will prompt you all of the questions in the Excel doc (one at a time), and you can click "YES", "NO", "n/a", "back" or "skip".  It will then write to the appropriate cell in the XLS file and put your response (and username).

You can write a comment in the comment box before you click a button, and it will add this comment to your response in the appropriate cell in the Excel doc.

Once you get to the end, click the "Create Summary" button (at the bottom right of the GUI), if you'd like it to create a summary in the form of a text file including all of your responses with comments.

**You must put a comment with any answer if you'd like to it end up in the summary.**


