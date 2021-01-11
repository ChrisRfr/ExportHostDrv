# ExportHostDrv
 Export Host Third-Party (OEM) Driver Packages

**Usage :**<br>
 ExportHostDrv.exe <Path_to_the_Destination_Folder><br>
 Example: ExportHostDrv.exe D:\ExportHostDriver<br>
<br>
![Alt text](/ExportHostDrv.jpg?raw=true "Reult")
<br><br>
It is ~the same as : Dism /Online /Export-Driver /Destination:D:\destpath<br>
With the list of drivers sorted and ordered as in the device manager :+1:<br>
Not just the DriverStore folder as with dism: iigd_dch.inf_amd64_56663c64bec44963, not really speaking.<br>
And it is faster than Dism.<br>
<br>
Forum: https://www.purebasic.fr/english/viewtopic.php?f=12&t=76563 <br />
<br>
Enjoy :)
