# Inventory-Utilities
Workbook and Scripts to manage inventory, connections and attributes

Creating resources in CloudShell begins with collecting some basic meta-data about the resource.  How much depends on the data model being used.  This document explains the process using examples from one specific installation but the process and tools are adaptable to nearly all use cases.

Basic Process
The process is straight forward with just a few steps.   You however first must add any custom atributes to the resource models created by the 2nd Gen Shells.  This is a task during initial setup or again later if a new attribute is identified.  Then you are read to load resources.

1.	Collect meta-data about the resource into 3 worksheets in the Inventory workbook, which provides inputs for the remaining steps.
2.	Run a script to create the resource and any sub-resources
3.	Run a script to set specific attributes for the resources/sub-resources
4.	Run a script to create connections to the L2 fabric

Moroe documentation is in the word doc included in the project.
