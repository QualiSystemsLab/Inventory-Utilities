## Inventory Utilities ##
The purpose of this utility is to assist in common inventory tasks related to CloudShell.
It is driven by a python script and an associated Excel file.  

The Excel file contains a number of tabs, each associate with a different function, along with an instructions tab, and a configuration tab.
You must ensure that the _settings_ tab on the Excel file is filled out and correct prior to running the script.

The Excel file must reside in the same working folder as the script, and must be named 'Inventory.xlsx'

#### Settings Tab ####
There are a series of settings that need to be configured prior to running the script:
* cs_host: The server hostname or IP of the Quali CloudShell Server being used
* cs_username: CloudShell Username to login with.  Any devices created will be attributed to this user
* cs_password: Password associated with the user
* cs_domain:  CloudShell Domain to log into; must be a domain available to the user
* cs_port:  CloudShell API port (Default is 8029, only change if installation is configured differently)
* logfilename:  Name for the log file to be created (or appended to) for this operation.  Log file is created in the local directory
* log level: Log Severity Level.  DEBUG by default, will capture each step.

#### Executing ####
In your command line, navigate to the folder and run the python script: 'python inventory_utilities.py'

You will be presented with 6 options, plus the option to exit:

1. Create and Autoload
2. Set Attributes
3. Set Connections
4. Bulk Load (1, 2, 3)
5. Add Custom Attributes
6. List Connections


## Options ##

#### Create and Autoload ####
The primary purpose of this function is to create a new item into inventory with in CloudShell.

The _Create and Autoload_ option will use the '1-CreateAndAutoLoad' tab in the Excel file.

The following columns are used in this tab:
* Ignore: if marked with __'Y' or 'y'__ will skip this row
* AutoLoad: if marked with __'Y' or 'y'__ will execute the Shell's _Autoload_ function after creation
* Parent: if the item is a child resource (e.g. card or port), this is the name of the it's parent
* Name: Name of this device
* Resource Family: the name of the Resource Family this item is modeled upon
* Resource Model: the name of the Resource Model (a subset of Family) that this item is modeled upon
* Domains: Enter the name of the domain(s) to which this unit is suppose to belong - comma separated list
* Address: Address of this resource
* Folder Path: To which folder with in CloudShell (resource manager) the device should be placed when created
* Connection Type: CLI Connection Type:  Auto, Console, SSH, Telnet
* User: Admin Username to use on the device
* Password: Password associated with the Username
* Enable Password: The OS Enable Password for the device
* Description: Device Description (if any)
* Driver Name: Name of the Driver to associate with this unit (for most shells there is only 1)
* SNMP Version: The SNMP Version the unit is configured for: v1, v2c, v3
* SNMP Read String: The SNMP Community Read String
* Location: Location information on the device (Lab, rack/row, etc.)

#### Set Attributes ####
The primary purpose of this function is to set attribute values associated with a device, attributes not associated with the autoload.  
This allows for the complete set of details to be set on the device when added to inventory.  
Can also be used to update large sets of attribute values, such as updating 'enable password'.

The _Set Attributes_ option will use the '2-SetAttributes' tab in the Excel file.

The following columns are used in this tab:
* Ignore: if marked with __'Y' or 'y'__ will skip this row
* ResourceName: Name of the device being modified/updated
* _AttributeName_: Set the name of the Attribute you want to update in the column header
    * Only need the basic name of the attribute, so for 2nd Gen Shells use _'Password'_ and __not__ _'WECSwitchShell.Password'_
    * Repeat as many times as needed.  If a resource does not have that attribute, leave blank.
    * This will not add an attribute that does not exist already on the Shell - for that use the 'Add Attribute' option

#### Set Connections ####
The primary purpose of this function is to set the connection map on a devices' sub-resource (Ports).

The _Set Connections_ option will use the '3-SetConnections' tab in the Excel file.

This allows for an easy way to set or update existing connections on a device.  
You should avoid putting in non-connectible device names into the Excel File.  Ideally it should just be the connectable device names.

The following columns are used in this tab:
* Ignore: if marked with __'Y' or 'y'__ will skip this row
* From: One side of the Connection
* To: The other side of the Connection
    * Leave blank to remove an existing connection
    
#### Bulk Load ####
This function will automatically run the following options in order:
1. Create and AutoLoad
2. Set Attributes
3. Set Connections

This is designed to be an 'Easy Button' to preform the 3 major actions needed to completely load a new device into CloudShell.

#### Add Custom Attributes ####
The primary purpose of this function is to add new attributes to 2nd Gen Shells.  
Currently there are two ways to do this, either via API (as done here), or to modify the _'shell-definition.yaml'_ file.

The _Add Custom Attributes_ option will use the '0-AddCustomAttributes' tab in the Excel file.

If any of these attributes are included used by the 'Set Attributes' option, this needs to be ran first.

The following columns are used in this tab:
* Ignore: if marked with __'Y' or 'y'__ will skip this row
* ModelName: Name of the Shell Model (Shell Name) to which to add the new Attribute to
* AttributeName: The name for the Attribute being added.
* DefaultValue: If you want to include a default value, do so here

#### List Connections ####
The Primary purpose of this function is to examine and capture all of a devices sub-modules (blades, ports, etc) and report what it's current connection is.

The _List Connections_ option will use the '4-ListConnections' tab in the Excel file.

This function will generate (or overwrite) a file named _current_connection.csv_ which will list out for each device the entire structure in the first column, and any mapped connections in the second.
Ideal use would be to then copy & paste the child names needed into the 'SetConnections' tab.
This ensures that all devices names are correct, and generally speeds the process.

The following columns are used in this tab:
* Device Names: List the name of the device you wish to list all children for