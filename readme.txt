IconWrks sample.
This is a 32 bit version of the iconwrks sample which originally
shipped with VB 1.0.  This project will also build the 16-bit version
in VB 4.0, but the 16-bit version may have problems running correctly
on Windows NT.  This is a straight port, with just enough changes made
to enable this project to run under 32 bit operating systems.

1) To build this application sample, either double click on the IconWrks.vbp icon 
under Windows (any version), or select this file from the File/Open menu in Visual 
Basic.  Once the application is loaded, either select Run/start to run
the program, or File/Make Exe to build an executable of the application.

1a) This code shows how to dynamically create a picture object in VB 4.0
by using windows bitmap handles and OLE API calls.  The process of creating
icons 'on the fly' is significantly different than under 16-bit OS's, where
VB was able to write to an icons memory image directly.

2)
|SAMPLE CODE.  Microsoft grants to you a royalty-free right to
|use and modify the source code version and to reproduce and
|distribute the object code version of the sample code, icons,
|cursors, and bitmaps provided within the Sample Code
|bin/folder on the SOFTWARE ("Sample Code") provided that you:
|(a) distribute the Sample Code only in conjunction with and as
|a part of your software product; (b) do not use Microsoft's
|name, logo, or trademarks to market your software product; and
|(c) agree to indemnify, hold harmless, and defend Microsoft
|and its suppliers from and against any claims or lawsuits,
|including attorneys' fees, that arise or result from your
|distribution of your software product.
|
|REDISTRIBUTABLE COMPONENTS.  Microsoft grants you a
|non-exclusive royalty-free right to reproduce and distribute
|the .DLL files included as part of the Sample Code provided
|that you: (a) distribute the .DLL files only in conjunction
|with and as a part of your software product; (b) do not use
|Microsoft's name, logo, or trademarks to market your software
|product; (c) agree to indemnify, hold harmless, and defend
|Microsoft and its suppliers from and against any claims or
|lawsuits, including attorneys' fees, that arise or result from
|the use or distribution of your software product; and (d)
|otherwise comply with the terms of this license agreement.

3) Place in Unsupported Samples directory.

4) Matthew J. Curland (MattCur)

5) VB4 16 & 32-bit editions, under Windows 3.1, Windows 95 or 
   Windows NT 3.51.

