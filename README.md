# KeepOutlookRunning

Little add-on to keep Outlook running when clicking on the exit button

Credit goes to @Tim Eck on Superuser (https://superuser.com/users/78284/tim-eck)
I'm just reuploading it on GitHub to back it up, this add-on works for Outlook 2010, 2013, 2016, 2019 and Outlook 365.

# Installation 

1. Install Microsoft Visual C++ 2010 SP1 Redistributable Package [32 bit](https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe) and [64 bit](https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe) (both required for 64 bit Windows)
2. Download [KeepOutlookRunning.dll](http://sourceforge.net/projects/keepoutlook/files/0.0.1/) 32 bit or 64 bit
3. Start Outlook as administrator (right click on it in Start Menu)
4. Go to File -> Options -> Add-Ins
5. At the bottom: Manage [COM Add-ins] press [Go...]
6. [Add...] the KeepOutlookRunning.dll file downloaded in step 2
7. Restart Outlook as a normal user

To exit Outlook after installing the add-on use File -> Exit

PS: In case the sourceforge.net link goes down, I add the .dll file to the project so you can get it from here
