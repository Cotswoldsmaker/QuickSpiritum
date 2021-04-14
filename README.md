# QuickSpiritum
A single digital solution to automate and speed up clinical software systems used in everyday clinical life.

## Overview
Quick Spiritum (QS) is built within the robotic process automation (RPA) open source program AutoHotKey. Quick Spiritum is a spin off from work I did on a database called Spiritum, which also used RPA. Spiritum is Latin for 'breath' (as I am from a respiratory background). Its main functionality is to automate and speed up daily clinical tasks. Systems that it can currently enter credentials for, control and search for patient results include:

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

## Not just screen scrapping...
QS also uses Microsoft's COMs, memory mapped files, system messages and REST APIs to communicate and control other systems.

## Installation
* Install AutoHotKey onto any computer you want to use QS on.
* Have a shared drive that users can access. If you do not have a shared drive, you should be able to use a stand alone version of QS soon. I am building this for other departments in our hospital.
* Basically clone this repro.
* Run the file '.\library\Private variables - RUN ME ONCE!.ahk'. This will create a link file and then a file outside of the git folder that you can safely store secret/trust/surgery information in that wont be uploaded to GitHub.
* Fill in the new 'Private variables.ahk' file (not the link file in the '\Library' folder) with information needed.
* You will need word templates with bookmarks for QS to fill in. Place in a QuickSpiritum\Templates\ folder. I will create some sample ones soon
* You will need copies of signatures for all users to be able to send requests. Place in QuickSpiritum\Signatures\ folder. I will create some sample ones soon
* When you want to allow others to use QS and install it on the desktop, run the 'Update Master File.ahk' script to move QS out of the Dev environment. Any user can then click on this master file to install it to their desktop.


## Notes
* QS turns off Imprivata from automatically entering credentials
* QS places a shortcut in the startup folder, so QS will start every time your OS restarts
* Has a quality improvement program (QIP) built directly into it. You run the QIP and it stores the time it takes to complete difference tasks both manually and then with the aid of automation. You can then compare the times and see the time saved.


## More to come
* I am still working actively on this project, and plan to expand the README file to make it easier to understand this program.


## Help needed
* I would greatly appreciate input from others to help improve the Quick Spiritum program. Also, if you find this program useful and use it at your trust/surgery, I would be greatful if you could run the QIP functionality and send me the results. I am hoping to get as much data as possible on how this program speeds up people's day. So far we have seen an average of a 29.7% reduction in the time it takes complete routine clinical tasks in our department.
* Anyone interested in building a function to control the new online NHS mail server (Office 365) in AutoHotKey. To basically open and also send emails with attachments? Messsage me if you are.
