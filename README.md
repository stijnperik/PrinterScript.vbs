# PrinterScript.vbs

------------------------------------------------------------------------------------
Over Complicated Printer Logon script
By stijnperik (2023)

USAGE:
- Set this as a user login script by group policy.
- Change the settings section to match your setup.
- For every printer that you have on the print server, create a new security group with the same name as the printer.
- Add users & computers to the printer groups on active directory.
- After login, network printers will be deleted and all relevant printers will be mapped.
- The default printer will default to the last known default printer.

- Additonaly you have the option of providing printers to all the computers with in a set ou group.
- To do this, create a security group, named with a set prefix, inside the ou. 
- Add the printer groups to this new security group.
		
This script will also inherit printers from sub group memberships.
		
VER: 0.9
-----------------------------------------------------------------------------------
