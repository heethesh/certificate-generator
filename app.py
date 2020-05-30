__author__ = "Heethesh Vhavle"
__copyright__ = "Copyright 2017, Heethesh Vhavle"
__license__ = "GPL"
__version__ = "1.0.2"
__maintainer__ = "Heethesh Vhavle"
__status__ = "Not Maintained"


import sys
from os import getcwd
from os.path import splitext, isfile, join, abspath
from urllib2 import urlopen, URLError

from Tkinter import Tk, Frame, Label, Entry, Button, OptionMenu, Checkbutton, IntVar, StringVar
from Tkinter import DoubleVar, TOP, DISABLED, LEFT, RIGHT, X, END
import tkMessageBox
from tkFileDialog import askopenfilename, askdirectory
from ttk import Progressbar

from PIL import Image, ImageFont, ImageDraw, ImageTk
from xlrd import open_workbook
from yagmail import SMTP, inline


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = getattr(sys, '_MEIPASS', getcwd())
        print base_path
    except Exception:
        print 'abspath'
        base_path = abspath(".")
    return join(base_path, relative_path)


FROM_EMAIL = ''  # SET THIS
FROM_EMAIL_PASSWORD = ''  # SET THIS

#-- RESOURCES --#
# NOTE: Update the paths in app.spec

# SET THIS RESOURCE PATH (maybe or maybe not?)
# I actually don't remember what this has to be set to and what it is!
# This is probably where the resources should go to.
# Maybe just leave this blank as it is and try it out as the resources are
# bundled with the application in app.spec, so it might work.
# All this is only required if you want to deploy the GUI as an .exe file.
FROOT = ''

# Certificate template backgrounds
BLANK_LIGHT = resource_path(FROOT + 'blank-light.jpg')
BLANK_DARK = resource_path(FROOT + 'blank-dark.jpg')
BLANK_CUSTOM = resource_path(FROOT + 'blank-dark.jpg')

# Certificate fonts
FONTL = resource_path(FROOT + 'Roboto-Light.ttf')
FONTM = resource_path(FROOT + 'Roboto-Medium.ttf')

# This is the logo in the email
PSLOGO = resource_path(FROOT + 'pslogo.png')
# GUI app icon
PSICON = resource_path(FROOT + 'psicon.ico')

# Some globals
EXCEL = ''
LOGO_PATH = ''
OUTPUT_FOLDER = ''
ENABLE = [False, False, False]
ERROR_LOG = ''
PING_TEST = 'http://www.google.com'

# Text that goes on the certificate
text = [''] * 6
text[0] = 'This is to certify that'
text[1] = 'from'
text[2] = 'has participated in the event'
text[3] = 'during National Level Annual Tech'
text[4] = 'Symposium, Phase Shift 2017, at BMS College of Engineering,'
text[5] = 'Bengaluru on the 15th and 16th of September 2017.'


def print_certificate(name, college, event, logo, bgmode):
    # global text, OUTPUT_FOLDER

    img = Image.open(bgmode)
    draw = ImageDraw.Draw(img)
    fontl = ImageFont.truetype(FONTL, 72)
    fontm = ImageFont.truetype(FONTM, 72)

    # This is where the text is printed on the certificate
    for i in range(0, 5):
        out = ''
        posy = 1200 + (i * 128)

        if i == 0:
            w1, h1 = draw.textsize(text[0], font=fontl)
            w2, h2 = draw.textsize(' %s ' % name, font=fontm)
            w3, h3 = draw.textsize(text[1], font=fontl)

            posx = 1754 - int((w1 + w2 + w3) / 2)
            draw.text((posx, posy), text[0], (0, 0, 0), font=fontl)
            draw.text((posx + w1, posy), ' %s ' % name, (0, 0, 0), font=fontm)
            draw.text((posx + w1 + w2, posy), text[1], (0, 0, 0), font=fontl)

        elif i == 1:
            w1, h1 = draw.textsize(' %s ' % college, font=fontm)
            w2, h2 = draw.textsize(text[2], font=fontl)

            posx = 1754 - int((w1 + w2) / 2)
            draw.text((posx, posy), ' %s ' % college, (0, 0, 0), font=fontm)
            draw.text((posx + w1, posy), text[2], (0, 0, 0), font=fontl)

        elif i == 2:
            w1, h1 = draw.textsize(' %s ' % event, font=fontm)
            w2, h2 = draw.textsize(text[3], font=fontl)

            posx = 1754 - int((w1 + w2) / 2)
            draw.text((posx, posy), ' %s ' % event, (0, 0, 0), font=fontm)
            draw.text((posx + w1, posy), text[3], (0, 0, 0), font=fontl)

        else:
            out = text[i + 1]
            w, h = draw.textsize(out, font=fontl)
            posx = 1754 - int(w / 2)
            draw.text((posx, posy), out, (0, 0, 0), font=fontl)

    # This is where an optional company logo is printed on the certificate
    if logo and var1.get():
        basewidth = 460
        imgl = Image.open(logo)
        wpercent = (basewidth / float(imgl.size[0]))
        hsize = int((float(imgl.size[1]) * float(wpercent)))
        imgl = imgl.resize((basewidth,hsize), Image.ANTIALIAS)
        try: img.paste(imgl, (3028, 20), imgl)
        except ValueError: img.paste(imgl, (3028, 20))

    save_name = '/%s - %s [%s]' % (event, name, college)
    img.save(OUTPUT_FOLDER + save_name + '.jpg', format='JPEG', subsampling=0, quality=95)
    return OUTPUT_FOLDER + save_name + '.jpg'


def file_check():
    # global FROOT
    files = ['blank-light.jpg', 'blank-dark.jpg', 'Roboto-Light.ttf', 'Roboto-Medium.ttf']
    msg = 'The following required files could not be found. '
    msg += 'Ensure that these files exist at the proper location and restart the program.\n\n'
    flag = True
    for f in files:
        if not isfile(FROOT + f):
            msg += FROOT + f + '\n'
            flag = False
    if not flag:
        tkMessageBox.showerror('Missing Files', msg)
    return flag


def warn_ext(mode):
    print 'EXT Error'
    if not mode: tkMessageBox.showwarning('File Error', 'Only .xlsx files are supported.')
    else: tkMessageBox.showwarning('File Error', 'Only JPG, PNG and BMP image formats are supported.')


def enable_buttons(mode):
    if mode:
        b6.config(state='normal')
        b7.config(state='normal')
        b9.config(state='normal')
    else:
        b6.config(state='disabled')
        b7.config(state='disabled')
        b9.config(state='disabled')


def state_check():
    global ENABLE
    print ENABLE
    if (ENABLE[0] and (not var1.get()) and ENABLE[2]) or (ENABLE[0] and ENABLE[1] and ENABLE[2]):
        enable_buttons(1)
    else:
        enable_buttons(0)        


def browser(entry):
    global ENABLE, EXCEL, LOGO_PATH, BLANK_CUSTOM, cmsgvar
    filename = askopenfilename()
    print filename

    f, ext = splitext(filename)
    if entry == e1:
        if ext != '.xlsx':
            ENABLE[0] = False
            warn_ext(0)
            return
        else:
            ENABLE[0] = True
            EXCEL = filename
            wb = open_workbook(EXCEL)
            sheet = wb.sheet_by_index(0)
            cmsgvar.set('Completed 0/%d' % (int(sheet.nrows)))

    if (entry == e2 or entry == e4):
        if ext not in ['.jpg', '.jpeg', '.png', '.bmp']:
            if entry == e2: ENABLE[1] = False
            warn_ext(1)
            return
        else:
            if entry == e2:
                ENABLE[1] = True
                LOGO_PATH = filename
            elif entry == e4:
                BLANK_CUSTOM = filename

    entry.delete(0, END)
    entry.insert(0, filename)
    state_check()


def ask_folder():
    global ENABLE, OUTPUT_FOLDER
    folder = askdirectory()
    print folder
    e5.delete(0, END)
    e5.insert(0, folder)
    OUTPUT_FOLDER = folder
    ENABLE[2] = True
    state_check()


def cb_invoke():
    print var1.get()
    states = ['disabled', 'normal']
    e2.config(state=states[var1.get()])
    b2.config(state=states[var1.get()])
    state_check()


def om_invoke():
    print bgvar.get()
    states = ['disabled', 'normal']
    e4.config(state=states[bgvar.get() == 'Custom'])
    b4.config(state=states[bgvar.get() == 'Custom'])    


def quit(root):
    root.destroy()


def send_email(name, event, email, imgpath):
    # global PSLOGO, yag
    mailtext = 'Dear %s,\n\nThank you for participating in the event %s conducted during Phase Shift 2017 at BMSCE.\nPlease find your attached Participation e-Certificate. We look forward for your participation at Phase Shift 2018 as well!\n\nCheers!\nPhase Shift 2017 Team\n\n*** This is an automatically generated email, please do not reply to this message ***\nIn case of any discrepancy, please contact your Event Coordinator.\n\n' % (name, event)
    contents = [mailtext, inline(PSLOGO)]
    yag.send(email, 'Phase Shift 2017 - %s Participation Certificate' % event,
        contents, attachments=imgpath)
    root.update()


def generate(mode):
    global BLANK_LIGHT, BLANK_DARK, BLANK_CUSTOM, LOGO_PATH, EXCEL, progress_var, cmsgvar
    BG_DICT = {'Light': BLANK_LIGHT, 'Dark': BLANK_DARK, 'Custom': BLANK_CUSTOM}

    name = ''
    college = ''
    event = ''
    email = ''

    wb = open_workbook(EXCEL)
    sheet = wb.sheet_by_index(0)

    enable_buttons(0)

    # Parse the spreadsheet
    for i in range(0, sheet.nrows):
        name = str(sheet.cell(i, 0).value).strip()
        college = str(sheet.cell(i, 1).value).strip()
        event = str(sheet.cell(i, 2).value).strip()
        email = str(sheet.cell(i, 3).value).strip()

        if college in ['BMS', 'BMSCE']: college = 'B.M.S. College of Engineering'
        elif college in ['DSCE']: college = 'Dayananda Sagar College of Engineering'
        elif college in ['RVCE']: college = 'R.V. College of Engineering'

        print '%s | %s | %s | %s' % (name, college, event, email)
        path = print_certificate(name, college, event, LOGO_PATH, BG_DICT[bgvar.get()])

        if (mode == 0):
            cmsgvar.set('Completed 1/1')
            progress_var.set(100)
            break

        cmsgvar.set('Completed %d/%d' % (i + 1, int(sheet.nrows)))
        progress_var.set(int(((i + 1) / float(sheet.nrows)) * 100))
        progressbar.update_idletasks()
        lab7.update_idletasks()

        if (mode == 2):
            send_email(name, event, email, path)

    wb.release_resources()
    enable_buttons(1)


def button_trigger(mode):
    global ERROR_LOG, progress_var

    try:
        wb = open_workbook(EXCEL)
        sheet = wb.sheet_by_index(0)
        progress_var.set(0)
        cmsgvar.set('Completed 0/%d' % (int(sheet.nrows)))

        if (mode == 2) and not check_internet():
            tkMessageBox.showwarning('No Internet Connection', 'Please check your internet connection and retry.')
            return
        else:
            global yag
            yag = SMTP(FROM_EMAIL, FROM_EMAIL_PASSWORD)

        generate(mode)

    except Exception as err:
        ERROR_LOG += '>>> ' + str(err) + '\n\n'
        tkMessageBox.showwarning('Unexpected Error Occured', ERROR_LOG + 'Please report the error(s) to heethesh@gmail.com')


def check_internet():
    try:
        urlopen(PING_TEST, timeout=1)
        return True
    except URLError as err:
        return False


##---- TK GUI STUFF ----##

root = Tk()
root.title('Phase Shift 2017 e-Certificate Generator')
root.minsize(width=465, height=310)
root.maxsize(width=465, height=310)
root.iconbitmap(default=PSICON)

## Excel
row1 = Frame(root)
row1.pack(side=TOP, fill=X, padx=3, pady=3)
lab1 = Label(row1, width=14, text='Excel Sheet', anchor='w')
lab1.pack(side=LEFT)
e1 = Entry(row1, width=40)
e1.pack(side=LEFT)
b1 = Button(row1, width=13, text='Choose File', command=lambda root=root:browser(e1))
b1.pack(side=LEFT, padx=3, pady=3)

## Logo
row2 = Frame(root)
row2.pack(side=TOP, fill=X, padx=3, pady=3)
lab2 = Label(row2, width=14, text='Company Logo', anchor='w')
lab2.pack(side=LEFT)
var1 = IntVar(root)
c1 = Checkbutton(row2, variable=var1, command=lambda root=root:cb_invoke())
c1.pack(side=LEFT)
e2 = Entry(row2, width=35, state='disabled')
e2.pack(side=LEFT)
b2 = Button(row2, width=13, text='Choose Image', state='disabled', command=lambda root=root:browser(e2))
b2.pack(side=LEFT, padx=3, pady=3)

## BG Select
row3 = Frame(root)
row3.pack(side=TOP, fill=X, padx=3, pady=3)
lab3 = Label(row3, width=14, text='Certificate BG', anchor='w')
lab3.pack(side=LEFT)

bgvar = StringVar(root)
bgvar.set('Dark')
w1 = OptionMenu(row3, bgvar, 'Dark', 'Light', 'Custom',  command=lambda root=root:om_invoke())
w1.config(width=15)
w1.pack(side=LEFT)

## Certificate File
row4 = Frame(root)
row4.pack(side=TOP, fill=X, padx=3, pady=3)
lab4 = Label(row4, width=14, text='Custom BG', anchor='w')
lab4.pack(side=LEFT)
e4 = Entry(row4, width=40, state='disabled')
e4.pack(side=LEFT)
b4 = Button(row4, width=13, text='Choose Image', state='disabled', command=lambda root=root:browser(e4))
b4.pack(side=LEFT, padx=3, pady=3)

## Output Folder
row5 = Frame(root)
row5.pack(side=TOP, fill=X, padx=3, pady=3)
lab5 = Label(row5, width=14, text='Output Folder', anchor='w')
lab5.pack(side=LEFT)
e5 = Entry(row5, width=40)
e5.pack(side=LEFT)
b5 = Button(row5, width=13, text='Choose Folder', command=lambda root=root:ask_folder())
b5.pack(side=LEFT, padx=3, pady=3)

## Progress Bar
progress_var = DoubleVar()
cmsgvar = StringVar()
cmsgvar.set('Completed 0/0')
row8 = Frame(root)
row8.pack(side=TOP, fill=X, padx=3, pady=3)
lab6 = Label(row8, width=14, text='Progress', anchor='w')
lab6.pack(side=LEFT)
progressbar = Progressbar(row8, variable=progress_var, length=245, maximum=100)
progressbar.pack(side=LEFT, fill=X)
lab7 = Label(row8, width=14, textvariable=cmsgvar, anchor='w')
lab7.pack(side=RIGHT, fill=X)

## Buttons
row6 = Frame(root)
row6.pack(side=TOP, fill=X, padx=7, pady=15)
b9 = Button(row6, text='Generate + Email', width=15, state='disabled', command=lambda root=root:button_trigger(2))
b9.pack(side=RIGHT, padx=3, pady=3)
b6 = Button(row6, text='Generate All', width=15, state='disabled', command=lambda root=root:button_trigger(1))
b6.pack(side=RIGHT, padx=3, pady=3)
b7 = Button(row6, text='Generate Sample', width=15, state='disabled', command=lambda root=root:button_trigger(0))
b7.pack(side=RIGHT, padx=3, pady=3)
b8 = Button(row6, text='Close', width=15, command=lambda root=root:quit(root))
b8.pack(side=LEFT, padx=3, pady=3)

## Info
row7 = Frame(root)
row7.pack(side=TOP, fill=X, padx=3, pady=3)
info = 'Phase Shift 2017 e-Certificate Generator by Heethesh Vhavle'
Label(row7, text=info, fg='#888888').pack()

# if not(file_check()): quit(root)
root.mainloop()
