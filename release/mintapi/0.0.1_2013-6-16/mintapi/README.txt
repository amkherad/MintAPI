Welcome to MintAPI ToolKit Library
===============

MintAPI is a ATL/OLE based framework. It consists of software libraries
and development tools.

MintAPI is a advanced ATL based library (Only Windows) provided for
Weak platforms like VB6 and includes many advanced and powerfull features
Such as threading ,process models and managements ,language globalization,
Advanced streaming ,advanced file io ,data management ,advanced data classes,
Plugin manager/loader ,advanced graphics and pixmap classes ,system registry,
Networks such as socket and more ,advanced MIDI engine ,NoteBuffer class,
Exception management ,including special syntaxes such as throw or out/inp(<</>>),
Automated application configuration management ,and more usefull and modern
Features , and more important ultimately pluginable...

MintAPI is developed as an open source project.

All information on MintAPI is available on the MintAPI Wiki on Sourceforge:
https://sourceforge.net/projects/mintapi/wiki/Home/

Overview
--------

You can use the MintAPI installation program to install the following components:

- MintAPI module (mintapi0.dll), Core library of MintAPI.
- MintAPI secondary layer (mintapi2ndlayer.dll), Basic MintAPI GUI tools.
- MintAPI.vb6.IDE.dll, VB6 IDE manager plugin.
- MintAPI proccess wrapper (mintapiwrapper.exe), Provides some advanced
  cross-proccess executions.
- MintAPI shell (shell.mintapi.exe), MintAPI shell !

Install MintAPI libraries to develop or run applications that need the MintAPI runtimes
or to try out example applications built with MintAPI.


Installing MintAPI
---------------

You can download MintAPI latest version from https://sourceforge.net/projects/mintapi/upload/release/.
The directory provides downloading
 full package of MintAPI,
 full package of MintAPI binaries,
 and MintAPI.dll separately.

Start the installation program like any executable on the development platform.

Select the components that you want to install and follow the instructions of
the installation program to complete the installation.


Directory Structure
-------------------

The default top-level installation directory is the directory "mintapi/<version>" in
your home directory, but you can specify another directory (<install_dir>).


Running Example Applications
----------------------------

You can open most example applications in the VB6 to build
and run them. Additional examples can be opened by browsing
<install_dir>/<version>/examples.

Building MintAPI from Source
-------------------------

See <install_dir>/<version>/src/HOW TO COMPILE and
for instructions on building MintAPI from source.


Developing MintAPI Applications
--------------------------

To develop a MintAPI application, you need to set up a project. You can install
MintAPI templates to create a project using these templates in VB6 IDE new project wizard.
Or if you have been installed the MintAPI IDE plugin then you can create new project
using it's wizard.

Also MintAPI IDE manager provides some features like creating language file - 
setting file and more... that you can link them to your project easily.

If you don't want to install MintAPI application templates, then you can create
your project in VB6 IDE then you need to make a refrence to MintAPI required
assemblies.

To include MintAPI library in other platforms like .net framework, there is some
limitations in execution of some features in MintAPI which described in
<install_dir>/<version>/src/Features And Limitations.txt.


Want to Know More?
-------------------

Much more information is available at:

- https://sourceforge.net/p/mintapi/wiki/informations


We hope you will enjoy using MintAPI!

- Ali Mousavi Kherad.
- Contact me at: alimousavikherad@gmail.com

- UNDER LGPL License