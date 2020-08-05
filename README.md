## PhaseShift 2017 e-Certificate Generator
Application which generates e-Certificates in bulk and emails them to participants. Developed for PhaseShift 2017.

**NOTE:** I no longer develop or maintain this repository. This was developed long back for a specific purpose and I only released the source as I got many requests. Use this only for reference. I've also included the Windows executable in `release` folder for reference, but this will not be able to send emails or provide the option to customize certificate content. **Please work with a Python developer and use the source code to customize it for your application, see instructions below.**

## Setup and Build

This application is developed in Python2 using Tkinter for GUI. I also provide the PyInstaller spec file so that you can use this to create an executable file for deployment. This application was developed for use on Windows, but should work on Linux as well I think.

The application basically takes an Excel sheet as an input, optional logo to display on top-right corner and prints some data on template certificate images in a batch. Finally, it helps email all the certificate images in a batch.

The background template images, fonts, some logos and icons are all bundled together using PyInstaller when building the executable, so make sure you update the paths in `app.py` and `app.spec` accordingly if you want to deploy. If you just want to run the script directly, you may not need to handle the resources in `app.spec`. 

**NOTE:** Please see `app.py` for comments and remember to set `FROM_EMAIL` and `FROM_EMAIL_PASSWORD`.

Refer docs folder for tutorial and sample Excel sheet.

## Application Tutorial
### Step 1

Prepare the excel sheet as per the format below.

**First Column** – Full Name (Capitalize first letter)  
**Second Column** – College Name (Full name preferred)  
**Third Column** – Event Name (Capitalize appropriately)  
**Fourth Column** – Email ID

Save the excel sheet in ***.xlsx*** format.

### Step 2
Open the ***PhaseShift 2017 e-Certificate Generator*** software and browse for your excel sheet.

### Step 3
You can select a company logo to be displayed on the top right corner. This is optional. Click the checkbox and select the company logo. **NOTE:** Make sure the logo is cropped properly and excess space is trimmed. **PNG** format is preferred.

### Step 4
You have the option to choose between a light and a dark background for the company logo (based on the logo colors). You can also create and choose a custom background template.

### Step 5
Select the output folder.

### Step 6
Generate your e-Certificates after preview!

**Generate Sample** – Creates a single sample certificate in the specified output folder (Preview).  
**Generate All** – Generates all the certificates to the output folder (Verify all once before email).  
**Generate + Email** – Generates all the certificates and emails them. 

**NOTE:**  For auto-mailing, a good internet connection is recommended as around 1.5 MB of attachments is to be uploaded per email. The software may sometimes freeze or hang, wait until the progress bar completes. Mailed ID will be **phaseshift.event@gmail.com**.
