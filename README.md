# CUCM Device Manager.exe

CUCM Device Manager Powershell executable for removing user devices in Cisco Unified Call Manager (CUCM) from stale/disabled users, doing so will remove the License they are holding within CUCM. Utilizes native Rest API calls to the CUCM server and only requires to be run on a system that has the Powershell Active Directory Module installed with no other dependancies. Since we are querying AD for user account information, you will want to run this tool from an account that has at least 'Read' access to the majority of OUs and Objects in your Active Directory Domain.

Now with more UI!

# Install

No install required, simply download the 'CUCM Device Manager.exe' for the executable or 'CUCM-Device-Manager-GUI.ps1' and run either on a machine with the Active Directory Powershell module installed. The function of both are identical, the exe is just packaged for convenience.

# Operation

Most of the operation is self explanatory, but here is a run down of the steps to remove a device from a user:

<p align="center"><img width="698" alt="ServerName-Login" src="https://user-images.githubusercontent.com/66786161/235771832-5a0a02ab-a07f-4ba8-b6e4-5021435b0c8a.PNG"></p>

1. Verify the local CUCM server FQDN is correct for your environment (shown above). Upon logging in successfully this servername will be saved in your local profile on the device you run it from. This should populate the saved servername upon re-running the tool. If there are any unknown errors running the tool verify the CUCM server version and change it accordingly, but the '12.5' value should suffice in most instances regardless of version.

<p align="center"><img width="691" alt="Login-Prompt" src="https://user-images.githubusercontent.com/66786161/235771881-46252fff-17de-452f-81aa-1959965f72ab.PNG"></p>

2. Upon clicking login you will be prompted to enter the server credentials, it is recommended to use the built in 'Admin' account as it will have rights to read/edit the CUCM database via the SOAP API this tool utilizes. Other accounts can be used but may need to have permissions edited in order for the tool to function properly.

Information on adding these permissions to an account in CUCM can be found here: https://www.cisco.com/c/en/us/td/docs/voice_ip_comm/cucm/admin/11_5_1/sysConfig/CUCM_BK_SE5DAF88_00_cucm-system-configuration-guide-1151/CUCM_BK_SE5DAF88_00_cucm-system-configuration-guide-1151_chapter_0100000.html

<p align="center"><img width="698" alt="Date-Selection" src="https://user-images.githubusercontent.com/66786161/235771930-1f98dc8c-3eaa-4193-ac17-a4c5fae6e3f6.PNG"></p>

3. Choose a date to look for accounts that have not logged in before. The tool defaults this date to 90 days prior to the current date. There are also buttons below the drop down that will change the selected date accordingly.

For example, in the above image the date January 18th, 2023 is selected. When we search for users in a future step, the tool will check Active Directory for user accounts that have not logged in since that date. Adjust this date to your desired search.

<p align="center"><img width="698" alt="Userlist" src="https://user-images.githubusercontent.com/66786161/235771976-026202c3-bbe0-4dcf-ad82-cb5732e0672e.PNG"></p>

4. The figure above shows the output of clicking 'Find Users' based only on the desired last logon date selected in the previous step. It takes the users found in Active Directory with the selected search options and finds users that currently have a device in CUCM/Jabber. This will then populate the list at the bottom of the form and provide information for each user found.

<p align="center"><img width="698" alt="Select-DisabledOU" src="https://user-images.githubusercontent.com/66786161/235772038-755335d7-d939-494a-ba86-98e1dc434d53.PNG"></p>

5. If desired you can expand the search to include users in the discovered Disabled OU in your Active Directory structure. Since there may be more than one, select from the drop down the desired OU search location.

<p align="center"><img width="698" alt="Options" src="https://user-images.githubusercontent.com/66786161/235772079-19e24d60-4305-4862-9d91-60e1d0b046e7.PNG"></p>

6. There is also an option to include all discovered Disabled Users in case there are ones that have not been moved to a Disabled OU.

<p align="center"><img width="698" alt="Select-Remove" src="https://user-images.githubusercontent.com/66786161/235772117-09384ae7-aff3-465f-8d8d-7a03df9bc996.PNG"></p>

7. Upon selecting the desired Date and Options, review the generated user list and determine which users you wish to remove Jabber/CUCM devices from. Check the box next to these users and click the 'Remove Devices' button.

<p align="center"><img width="698" alt="DevicesRemoved-ConsoleOutput" src="https://user-images.githubusercontent.com/66786161/235772161-affa42c6-7d0d-41e6-ad0e-7a191609307b.PNG"></p>

8. After clicking 'Remove Devices' you will start to see output in the console box showing the tasks to tool is conducting. Most success/error/failure output will be redirected to this console box. Upon completing the tasks for each selected user the tool will refresh the user list as final verification of removed devices. From here if desired you can copy/paste the console output to a text file to conduct any required Maintenance Logs.

# Customization

I embedded my Icon in the source file as a base64 image (line 29), if you wish to change the icon just swap out the string there and rebuild the exe for your purposes using a complier like [ps2exe](https://github.com/MScholtes/PS2EXE). You can also edit the default server name in the source file (line 10) before compiling.

Submit any feature requests and feedback, the methods I used here can be leveraged for a large number of CUCM tasks, not just removing user devices, this was just a specific need and use case I had in the environment I work in to free up CUCM licenses that stale users in AD were holding.
