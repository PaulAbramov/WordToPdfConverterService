# WordToPdfConverterService
Service for Windows to automatically convert word documents from a directory to pdf

I used C# for this, so it should work with the simple build command in Visual Studio.
This builds the .exe file, which is also included in the compiled solution.

using the "Visual Studio 2019 Developer Command Prompt" navigate to the build folder and install the service with "Installutil WordToPdfConverter.exe"
open "services" by going on windows start and typing services.
Look for "Word to PDF Converter"
Double click it, so the popup opens, where you can enter startparameter.
You can try and enter the folderpath where the .docx files will be stored in the startparameter.
It worked 2 times for me, then I made some changes and it did not work anymore.
The default startparameter is "C:\Users\Administrator\Desktop\WordDocuments"
start the service.

All logs can be seen in the "Event Viewer", which can be opened by going on windows start and typing in "event viewer".
Expand "Applications and Services Logs" on the left side.
The subentry "DocxTOPdfConverterLog" is what we are looking for.
click on it and it shows all logentries.

Now you should see if the correct path is used or not.

If it is not the correct path you can change the default path as follows:

shutdown the service in the "services" App

windows start -> enter "regedit" and open it.

In regedit "Computer\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\DocxToPdfConverterService" change Imagepath 2nd parameter to the path of the folder with the .docx files

start the service in the "servies" App, now it should use the correct path to the folder with the .docx files.


Troubleshooting:

If it is not working, add following directories:
C:\Windows\System32\config\systemprofile\Desktop
C:\Windows\SysWOW64\config\systemprofile\Desktop

In the "services" app double click on the Service "Word to PDF Converter", in the Popup go to the tab "Logon" and set the checkbox to allow exchange between the service and the desktop.
