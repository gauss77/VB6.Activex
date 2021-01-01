# Zip and Unzip Using VB5 or VB6

![Capa](/capture.png?raw=true "Capa")

Author: Chris Eastwood
Author 's WebSite: http://www.eastwood60.freeserve.co.uk/

About Zip32.DLL And Unzip32.DLL
This code shows how to use the freeware InfoZip Zip32.DLL and Unzip32.DLL files from the http://www.cdrom.com/pub/infozip/ website.
The InfoZip DLL's are open-source DLL's that are available for programmers to utilise free of charge. They are standard C DLL's and were very tricky / impossible to interface with VB until VB5/6 offered the 'addressof' operator.

# About this Code
             
I based this code on the Zip/Unzip examples that come with the downloadable DLL's and encapsulated that code into re-usable class modules that can either be used in an ActiveX DLL, or just placed into your project as required. The downloadable project contains the source-code for an ActiveX DLL called CGZipLibrary and a sample project which uses that DLL.
The example program simply takes everything in it's own directory and zip's it into a file. You can then unzip this file straight into the TEMP directory. You can check the structure of the Zip file using WinZip (or similar) if you want to check that the Zip is intact (I've had no problems with it at all).
The Class Modules CGZipFiles and CGUnzipFiles have a lot of properties - it's worth reading through the code in these classes to learn how to get the best from them. A sample piece of code using the Zip class could be as simple as :

```
Dim oZip As CGZipFiles

Set oZip = New CGZipFiles

oZip.ZipFileName = "\MyZip.Zip"
oZip.AddFile "c:\mystuff\myfiles\*.*"
oZip.AddFile "c:\mystuff\mymedia\*.wav"

If oZip.MakeZipFile <> 0 Then
   MsgBox oZip.GetLastMessage
End If

Set oZip = Nothing
```

#The code for Unzipping files is just as straight-forward :
```
Dim oUnZip As CGUnzipFiles

Set oUnZip = New CGUnzipFiles

oUnZip.ZipFileName = "\MyZip.Zip"
oUnZip.ExtractDir = App.Path ' or whatever you like !
oUnZip.HonourDirectories = True ' keep dir structure
If oUnZip.Unzip <> 0 Then
  MsgBox oUnZip.GetLastMessage
End If

Set oUnZip = Nothing
```
Note : This code is provided as-is - I know that there are a few bugs in the interface to the Zip/Unzip DLL's (including a frustrating one that's got a real cludge work around - see if you can spot it). It's designed as a starting point to allow you to use Zip and Unzip functionality in your programs. You should really place the Zip32.DLL and Unzip32.DLL into your \windows\system directory (although you can just have them in the same directory as the ActiveX DLL). I've been a version of this code for several weeks in a production application and it's just been superb. I hope you enjoy the code !
