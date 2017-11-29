General
=======

Protection
----------

Most files are protected. Not because I do want to hide my stuff, but to prevent accidental edits, as it voids the digital signature. If you want to have a look at the code or change the sheets, use **1234** as password for unlocking

Also when downloading these excel files you might need to right click it first, select "properties" and the on the lower right, click on "allow" if it is present

ExcelPlayer
===========

This is an Excel sheet, that plays Video and Audio files you have the codec for. Useful if at work and you are monitored, as the playback (even the video window) counts as excel usage.

Utilizes macros.

How To
------

Open the File and confirm the security bar on top.
Go into the playlist sheet and add files. The easiest way to accomplish this, is to right click on a media file on your computer while holding shift. Select "copy as path" from the menu, than paste the path into excel. **For now, remove the quotes** of the file.

Version
-------

The File has been tested under Windows 7 (64 bit) using Office 2010 (32 bit).

Excplorer
=========

A mixture between Excel and Explorer

A simple Internet Explorer rendering in Excel to browse the web while it is still counted towards activities in Excel.

Utilizes macros.

How To
------

Go to sheet 3 and click on the Start button. You can only have one window open at a time, but when closing it you can click this button again.

Version
-------

The File has been tested under Windows 7 (64 bit) using Office 2010 (32 bit).

CMD
===

This file copies the CMD from the windows\system32 directory onto your desktop and overwrites the portion that admins use to prevent people from running it.
This allows you to run cmd.exe on restricted systems.
It also has a function for patching regedit.exe in it, but for some reason it doesn't works.

Utilizes Macros

How To
------

Open Workbook, Push the appropriate button

Version
-------

The File has been tested under Windows 7 (64 bit) using Office 2010 (32 bit).

MazeGen
=======

A Maze generator and solving assistant

Allows you to generate and solve mazes

Utilizes macros.

How To
------

Allow Macros, then you should see a Tab "Add-Ins" in the ribbon with 3 new buttons.
Those are in the order from left to right:

- Generate new maze
- Reset solving progress
- Clear document

Click on Clear document, then on "Generate maze".
Enter width and height of the maze.
The maze generates now. If you enter big values (>150),
excel may freeze up. Just wait for it to finish to become responsive again.

After the maze is generated, use the arrow keys to navigate through the maze.
The computer will autmatically move as far until another decision is needed
(crossing, corner or dead end)

You can reset the progress at anytime with the reset button.

You cannot save the maze at the moment.
You can save the document, but you cannot continue to solve the maze.
The reset button should however allow you to solve the already present maze again.

Version
-------
The File has been tested under Windows 7 (64 bit) using Office 2010 (32 bit).

ExcelDecrypt
============

Not an excel file.
This file decrypts an office 2003 file.
You can use it for word and powerpoint documents as well.
Source file and a precompiled and signed exe file are present in the repository.

You can also use [My online service](http://home.ayra.ch/unlock/).
The online service also supports Office 2010 documents

How To
------

This is a command line tool:

    excelDecrypt.exe <infile> <outfile>

Version
-------

The File has been tested under Windows 7 (64 bit)
Since it is a basic C source code file, without any windows specific code, it will also run under DOS or linux if compiled.
