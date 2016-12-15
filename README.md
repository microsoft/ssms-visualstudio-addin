# ssms-visualstudio-addin

This is a demo to show how to build an add-in that can be used from within teh Visual studio tooling that we 
provide for Dynamics for Operations developers.


The demo is intended to be instructive, but also useful. It basically allows you to work with Dynamics 
tabular objects in the Microsoft SQL Server Management Studio (SSMS) tool. It will harvest interesting metadata for
the tabular object selected in the model view, and open a window in SSMS allowing you to do advanced queries against
the data. The tool handles table inheritance and relations to other tables, generating the joins as required.


With this tool you can: 

1.	Select table root node, showing all the data from the table. 
2.	Select one or more table fields, shows only the data for the selected field(s).
3.	Select one or more table field group, showing data for the fields in the selected field group(s). 
4.	Select one or more table relations, generating the join on the criteria mentioned in the relation(s).

Once the window has been opened in SSMS you can edit as required for your purposes.

## Building the tool
The code should be easy to build in Visual Studio once you fix up the references to
point to the place where your product specific assemblies live. The output from a 
successful compilation is an assembly that can subsequently be deployed using the
encloded batch file.

_Note_: You will need to restart Visual Studio to get see the new add-in.

## Using the tool
Once the tool is build and deployed as described above, it will be available in
the add-ins menu in the Dynamics 365 menu. This add-in is designed to be applied 
on a selected designer node, so you have to open the table in the designer to 
use it; you cannot use it directly from the Application Explorer at this time.
 


