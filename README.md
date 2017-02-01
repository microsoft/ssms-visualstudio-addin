# ssms-visualstudio-addin

This is a demo to show how to build an addin that can be used from within the Visual studio tooling that we 
provide for Dynamics 365 for Operations developers.


The addin is intended to be instructive, but also useful. It basically allows you to work with Dynamics 
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
batch file that is part of the project. Currently the command file is run when the 
compilation is completed. If this is not what you want, you can change the post-build
event command line in the Build Events tab in the project properties.

_Note_: You will need to restart Visual Studio to get access to the new add-in.

## Using the tool
Once the tool is built and deployed as described above, it is available in
the addins menu in the Dynamics 365 menu. This particular addin is designed to be applied 
on a selected designer node, so you have to open the table in the designer to 
use it; you cannot use it directly from the Application Explorer at this time.

## Creating Add-ins of your own
It is easy to get started writing your own addins. Just open Visual Studio, create a new project 
and choose the template called "Developer tools Addin" from the "Dynamics 365 for Operations" set.
This will create a project with examples of both a Mainmenu addin (i.e. an addin that is visible 
on the Addins menu on the Dynamics 365 menu, and that is not tied to any particular metadata artifact), 
and a Designer addin which is appears when a particular artifact is selected in the designer. There are 
TODO comments indicating where you should put your code.


After you have successfully compiled you solution, you can invoke the InstallToVS batch file to 
deploy the assembly so that Visual Studio will pick it up the next time it starts.

This project has adopted the [Microsoft Open Source Code of
Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct
FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com)
with any additional questions or comments.




