# Jabber Device Manager.exe

Jabber Device Manager Powershell executable for removing user devices in Cisco Unified Call Manager (CUCM) from stale/disabled users, doing so will remove the License they are holding within CUCM. Utilizes native Rest API calls to the CUCM server and only requires to by run on a system that has the Powershell Active Directory Module installed with no other dependancies. Since we are querying AD for user account information, you will want to run this tool from an account that has at least 'Read' access to the majority of OUs and Objects in your Active Directory Domain.

Now with more UI!

# Install

No install required, simply download the 'Jabber Device Manager.exe' for the executable or 'Jabber-Device-Manager-GUI.ps1' and run either on a machine with the Active Directory Powershell module installed. The function of both are identical, the exe is just packaged for convenience.

# Operation

Most of the operation is self explanatory, but here is a run down of the steps to remove a device from a user:

<p align="center"><img width="698" alt="ServerName-Login" src="https://github.build.ge.com/storage/user/114690/files/8f80c13f-bcfc-4145-8aa4-7c29bade428e"></p>

1. Verify the local CUCM server FQDN is correct for your environment (shown above). Upon logging in successfully this servername will be saved in your local profile on the device you run it from. This should populate the saved servername upon re-running the tool. If there are any unknown errors running the tool verify the CUCM server version and change it accordingly, but the '12.5' value should suffice in most instances regardless of version.

<p align="center"><img width="691" alt="Login-Prompt" src="https://github.build.ge.com/storage/user/114690/files/637be4fd-a253-48fd-965c-73729f084423"></p>

2. Upon clicking login you will be prompted to enter the server credentials, it is recommended to use the built in 'Admin' account as it will have rights to read/edit the CUCM database via the SOAP API this tool utilizes. Other accounts can be used but may need to have permissions edited in order for the tool to function properly.

Information on adding these permissions to an account in CUCM can be found here: https://www.cisco.com/c/en/us/td/docs/voice_ip_comm/cucm/admin/11_5_1/sysConfig/CUCM_BK_SE5DAF88_00_cucm-system-configuration-guide-1151/CUCM_BK_SE5DAF88_00_cucm-system-configuration-guide-1151_chapter_0100000.html

<p align="center"><img width="698" alt="Date-Selection" src="https://github.build.ge.com/storage/user/114690/files/2b9924dd-74a6-4a8d-ad4f-db2c330237fc"></p>

3. Choose a date to look for accounts that have not logged in before. The tool defaults this date to 90 days prior to the current date. There are also buttons below the drop down that will change the selected date accordingly.

For example, in the above image the date January 18th, 2023 is selected. When we search for users in a future step, the tool will check Active Directory for user accounts that have not logged in since that date. Adjust this date to your desired search.

<p align="center"><img width="698" alt="Userlist" src="https://github.build.ge.com/storage/user/114690/files/1e9f14f6-a769-4614-a6c1-952b02714713"></p>

4. The figure above shows the output of clicking 'Find Users' based only on the desired last logon date selected in the previous step. It takes the users found in Active Directory with the selected search options and finds users that currently have a device in CUCM/Jabber. This will then populate the list at the bottom of the form and provide information for each user found.

<p align="center"><img width="698" alt="Select-DisabledOU" src="https://github.build.ge.com/storage/user/114690/files/988e0733-1a14-4bcb-90e3-9dd837cf8016"></p>

5. If desired you can expand the search to include users in the discovered Disabled OU in your Active Directory structure. Since there may be more than one, select from the drop down the desired OU search location.

<p align="center"><img width="698" alt="Options" src="https://github.build.ge.com/storage/user/114690/files/4251ec20-a46e-4244-86d7-1b84f0105eaa"></p>

6. There is also an option to include all discovered Disabled Users in case there are ones that have not been moved to a Disabled OU.

<p align="center"><img width="698" alt="Select-Remove" src="https://github.build.ge.com/storage/user/114690/files/670cbc3b-6de9-4a70-9542-2e36a16df309"></p>

7. Upon selecting the desired Date and Options, review the generated user list and determine which users you wish to remove Jabber/CUCM devices from. Check the box next to these users and click the 'Remove Devices' button.

<p align="center"><img width="698" alt="DevicesRemoved-ConsoleOutput" src="https://github.build.ge.com/storage/user/114690/files/8a4e8a51-3377-4429-b328-3a57d32db42a"></p>

8. After clicking 'Remove Devices' you will start to see output in the console box showing the tasks to tool is conducting. Most success/error/failure output will be redirected to this console box. Upon completing the tasks for each selected user the tool will refresh the user list as final verification of removed devices. From here if desired you can copy/paste the console output to a text file to conduct any required Maintenance Logs.

# Jabber-Device-Manager.ps1
(Powershell Console Script version semi-deprecated but still usable)

Powershell tool to facilitate management of Cisco Unified Call Manager (CUCM) users and devices via CUCM SOAP API using
native AXL Rest calls.

Must be ran on a device with the Active Directory Powershell module available and you must have admin login credentials
to the CUCM server, no other dependancies are required.

When ran it will check disabled users in AD and generate a list that have a Phone Device associated with them in CUCM. It
will then prompt if you would like to remove the associated devices for each user. Selecting to remove the device will
remove it from associated devices for the user and delete the device in CUCM, thus freeing up that license from the
available license pool.

Planned future featuers incude a UI and further user/device management within the tool.

_Jake Hillman_

_Belcan Contractor_

![114690](https://github.build.ge.com/storage/user/114690/files/6c719b95-25ed-426e-a542-a7e94d5d8f9e)
