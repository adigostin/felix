
This project can be used for debugging the VS package with the "Native Only"
debugger. The necessary project settings should already be configured in the
project settings, and this project should already be the startup project.

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