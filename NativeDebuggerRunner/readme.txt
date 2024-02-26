
This project can be used for debugging the VS package with the "Native Only"
debugger. To use it, you have to make these setting in Project Properties:
In Solution Explorer, right-click the _ReadMe-NativeDebuggerRunner and choose Properties.
The go to Debugging, select Configuration: All Configurations,
select Platform: All Platform (this one may be missing in VS 2019 - that's ok),
then select Debugger to Launch: Local Windows Debugger, and then set these properties:
 - Command           = $(VSAPPIDDIR)$(VSAPPIDNAME)
 - Command Arguments = /rootsuffix Exp
 - Debugger Type     = Native Only
Click OK to close the Project Properties dialog, then right-click
this project in Solution Explorer and choose Set as Startup project.
You are now ready to build it and debug it.

If you ever delete the .suo file from the .vs directory,
you will need to make these settings again.

When you start debugging, an experimental instance of Visual Studio opens.
In this experimental instance you can create a new project: just type "Z80"
in the drop down where you search for template keywords, and double-click
on Z80 Project. When you create or open a Z80 Project, the Simulator tool window
should open too; if it doesn't, open it yourself from Main Menu -> View ->
-> Other Windows -> Simulator.
Enjoy!


Additional info about the Native Debugger Runner project:

The VS package is built and deployed by the Z80PackageVSIX project,
but that project is a C# one, and starting debugging from C# launches
the "Managed Only" or "Mixed" debuggers. These are horribly slow.
In Project Dependencies, a dependency exists on the Z80PackageVSIX.
This dependency tells VS to always build and deploy the VSIX project first
when your press Start Debugging.