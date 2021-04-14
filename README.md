# QuickSpiritum
A single digital solution to automate and speed up clinical software systems used in everyday clinical life

## Overview
Quick Spiritum (QS) is built within the robotic process automation (RPA) open source program AutoHotKey. Its main functionality is to automate and speed up daily clinical tasks. Systems that it can currently enter credentials for, control and search for patient results include:

* Sunrise EPR
* Trakcare
* InfoFlex 
* PACS 
* Lab results program (called PAS) 
* DigiDictate 
* Find lung function results 
* Open a purpose built Microsoft Access database called Spiritum (QS is actually a spin off from the database) 
* Noxturnal sleep study software 
* ICE 

It can also create requests /referrals for:

* PET-CT
* Lung function
* Bronchoscopy
* Healthy lifestyles Gloucestershire
* Sleepstation

QS can also send
* SMS messages, emails or letters (via the GOV.uk Notify site) to patients. Needs to be smartened up a little and have a list of messages that you can send.

## Not just RPA
QS also uses Microsoft's COMs, memory mapped files, system messages and REST APIs to communicate and control other systems.

## Help
Anyone interested in building a function to control the new online NHS mail server (Office 365) in AutoHotKey. To basically open and also send emails with attachments?

## Installation
* Install AutoHotKey onto any computer you want to use QS on.
* Have a shared drive that users can access.
* Basically clone this repro.
* Run the file '.\library\Private variables - RUN ME ONCE!.ahk'. This will create a link file and then a file outside of the git folder that you can safely store secret/trust/surgery information in that wont be uploaded to GitHub.
* Fill in the new 'Private variables.ahk' file (not the link file in the '\Library' folder) with information needed.
* You will need word templates with bookmarks for QS to fill in. Place in a QuickSpiritum\Templates\ folder. 
* You will need copies of signatures for all users to be able to send requests. Place in QuickSpiritum\Signatures\ folder.
* When you want to allow others to use QS and install it on the desktop, run the 'Update Master File.ahk' script to move QS out of the Dev environment. Any user can then click on this master file to install it to their desktop.


## Notes
* QS turns off Imprivata from automatically entering credentials
* QS places a shortcut in the startup folder, so QS will start every time your OS restarts
* Has a quality improvement (QI) program built directly into it. You run the QI program and it stores the time it takes to complete difference tasks manually and then with the aid of automation. You can then compare the times and see the time saved. 
## More to come
* I am still working actively on this project, and plan to expand the README file to make it easier to understand this program.
