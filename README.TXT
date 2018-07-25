Welcome to MintAPI ToolKit Library
==================================

MintAPI is a COM/OLE based framework. It consists of software libraries
and development tools.

MintAPI is an advanced OLE based library (Only Windows) provided for
weak platforms like VB6 and includes many advanced and powerfull features
such as threading, IPC and process managements, globalization, collections,
advanced streaming, advanced file IO, data management, advanced data classes,
plugin manager/loader, advanced graphics and pixmap classes, system registry,
networks such as socket and more, advanced MIDI engine, NoteBuffer class,
Exception management, including special syntaxes such as throw or out/inp(<</>>),
automated application configuration management, and more usefull and modern
features;
And more important it's ultimately pluginable...

MintAPI is developed as an open source project.

All information on MintAPI is available on the MintAPI Wiki on Sourceforge:
https://sourceforge.net/projects/mintapi/wiki/Home/


Overview
---------

You can use the MintAPI installation program to install the following components:

- MintAPI module (mintapi0.dll), Core library of MintAPI.
- MintAPI runtime (mintruntime0.dll), MintAPI's Runtime library (code generators, reflections, ...).
- MintAPI secondary layer (mintapi2ndlayer.dll), Basic MintAPI GUI tools.
- MintAPI file formats (mintff0.dll), Provices file format support for library.
- MintAPI.vb6.IDE.dll, VB6 IDE manager plugin.
- MintAPI host service process (mintsvchost.exe), Host process for MintAPI services.
- MintAPI shell (shell.exe), MintAPI shell !

Install MintAPI libraries to develop or run applications that need the MintAPI runtimes
or to try out example applications built with MintAPI.


Installing MintAPI
-------------------

You can download MintAPI latest version from https://sourceforge.net/projects/mintapi/upload/release/.
The directory provides downloading
 full package of MintAPI,
 full package of MintAPI binaries,
 and MintAPI.dll separately.

Start the installation program like any executable on the development platform.

Select the components that you want to install and follow the instructions of
the installation program to complete the installation.


Directory Structure
--------------------

The default top-level installation directory is  "mintapi/<version>" in
your home directory, but you can specify another directory (<install_dir>/<version>).


Running Example Applications
-----------------------------

You can open most example applications in the VB6 to build
and run them. Additional examples can be opened by browsing
<install_dir>/<version>/examples.


Building MintAPI from Source
-----------------------------

See <install_dir>/<version>/src/HOW TO COMPILE to see
instructions on building MintAPI from source.


Developing MintAPI Applications
--------------------------------

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
<install_dir>/<version>/src/Features And Limitations.txt, see the changelog to
see if these limitations are proved.
- Also there is a third-party tool library called "dotnetwrapper" that let you 
  connect full featured mintapi to .net applications, check it in the installation
  to install on your platform.


Changelog
----------

- 1.0.0000  MintAPI transfered from VB6 to C++ and become Cross-Platform! [I WISHING...]

- 0.0.0700  Asyncronous Programming Model (APM)-liked structure added to asyncronous classes and methods.
- 0.0.0672  Asyncronous methods (like .net) take place in all IO and more classes.
- 0.0.0671  Establishing threading foundation, and asyncronous execution.
- 0.0.0651  Create a strong streaming foundation for framework.
- 0.0.0651  Begin to layering MintAPI and remove secondary technologies from core library like UI.
- 0.0.0620  Establishing some of .net framework methods like IObject class and so more...
- 0.0.0013  Creating a powerfull exception managing.
- 0.0.0009  Create a class streaming foundation like POSIX library for all classes in framework.


Want to Know More?
-------------------

Much more information is available at:

- https://sourceforge.net/p/mintapi/wiki/informations


We hope you will enjoy using MintAPI!

- Ali Mousavi Kherad, and other contributors.
- Contact me at: alimousavikherad@gmail.com / amgmail@users.sourceforge.net

- UNDER GPL/LGPL/FDL Licenses
- It also provided under a responsible and supported commercial license.