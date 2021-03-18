# Update-BillableTime
A PowerShell script to create Billable Time in Microsoft System Center Service Manager for various Work Item types using Outlook calendar appointments and meetings.

# Requisites
Outlook and the Service Manager console need to be installed on the workstation.

The script is designed to be used with a Service Manager console task, but can also be used in a PowerShell console.

It is not required to have Outlook running at the time of script execution, but if you do have Outlook open, do not close Outlook while the script is running as this will cause errors.

The script is dependent on the Syliance.BillableTime.Extension (paid version) which is available from ITnetX (https://www.itnetx.ch/) and expands Billable Time to all Work Items in Service Manager and not just Incidents, it also enables comments and other features that expand on the basic built in Billable Time features in Service Manager.
However, it may well be possible to adapt the script to be used with other Billable Time solutions for Service Manager.

# How to use
1. Download and store the script to a network location so that all analysts can access it from the console task.
2. Create a new Service Manager console task from the library node in the wunderbar.
3. Give the task a name and a description.
4. Because this script updates the Billable Time on any Work Item, I recommend that you choose the 'Work Item' target class.
5. Select an existing Management Pack to store the task or create a new Management Pack for this purpose.
6. Select 'Work Item Folder Tasks' under Categories.
7. On the Command line tab:
   * Full path to command: PowerShell
   * Parameters: `–ExecutionPolicy Bypass -File .\Update-BillableTime.ps1` - *'–ExecutionPolicy Bypass' will bypass the execution policy only for this script. This isn't needed if the script is signed.*
   * Working directory: *\\youynetworklocation\*
   * Show output when this task is run: Ticked 
8. Press OK to save the task.
9. Under the Administration node of the wunderbar, navigate to Security -> User Roles.
10. Update the appropriate user roles so that the task is viewable in the console to analysts by checking the name of the task under the tasks node.
11. Restart the Service Manager console and select a Work Item. You should see the name of your new task anyware in the Work Items node.
12. Click on the task to run it. You should see a console window that displays the progress of the script. [Like this](/images/GPPTasksCommon.JPG?raw=true "GPP Files common tab")

# More Information
Tested using Service Manager 2016 and 2019, Outlook 2016, 2019 and O365 
