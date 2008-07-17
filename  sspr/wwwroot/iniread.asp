<%
'http://cwashington.netreach.net/depo/view.asp?Index=553&ScriptType=vbscript
'Author: Nolan Bagadiogn
'Description:
'This is based on Antione Jean-Luc's 1998 script, but now uses Regular Expressions to find sections and keys.
'Benefits:
'  no longer case sensitive;
'  not white space sensitive;
'  adds two more functions:
'   - Can return array of sections in file.
'   - Can return array of keys under a section in a file.
'===================
' Revision History
' 1.0 27 Jul 2001 NB Original version from website post
' 1.1 12 May 2006 PJW Merged recommended changes from other posts and
'     add some other changes to make the script work
'     better.
' 1.2 30 Aug 2006 PJW Commented out the msgbox line in the WRITEINI
'     Sub, as it would prevent the script from running
'     autonomously (i.e. without user intervention).
' 1.3 06 Sep 2006 PJW Was getting error "Unexpected quantifier" when
'     run with VBScript 5.1. Changed
'     reSection(Any).Pattern = "^\s*\[(.*?)\]\s*"
'     to
'     reSection(Any).Pattern = "^\s*\[.+\]\s*"
' 1.4 23 Mar 2007 PJW WriteIni is creating temp_ini.ini, making the modifications in it,
'     deleting the original *.ini, and then renaming temp_ini.ini to the
'     filename of the original *.ini. Using this method loses any specific
'     permissions (not inherited) that were on the original *.ini.
'     Changed script so that the original *.ini is repopulated with the
'     new settings and temp_ini.ini is deleted instead. This maintains the
'     permissions on the original *.ini.
'    There was a mistake in reSection(Any).Pattern, causing it to not remember matches.
'     It was instead returning "$1" for every match.Value. Changed
'     reSection(Any).Pattern = "^\s*\[.+\]\s*"
'     to
'     reSection(Any).Pattern = "^\s*\[(.+)\]"
'     Note: also removed the trailing whitespace (\s*) as this was just confusing matters.
'    Found that when the trailing whitespace (\s*) was in the .Pattern,
'     tempSection=tempSection & reSection.Replace(line, "$1") & ","
'     would return more that the section name if there was characters after the closing ]
'     For example, if the line with the section name was
'     [Section Name];
'     it would return "Section Name;"
'     Changed this to use
'     For Each objMatch in colMatches...
'     instead. This returned [Section Name], so had to trim [ and ] off.
'

'START FUNCTIONS AND SUBS HERE
'===================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalSCRIPT(TM)
'
' NAME: VBScript ReadINI/WriteINI
'
' AUTHOR: Bagadiong, Nolan
' DATE : 7/27/2001
'
' COMMENT: Uses Regular Expressions to find keys and values; no longer case
' or space sensitive
'
' If you pass READINI(file, section, key), function returns scalar of the value
' assigned to key
'
' If you pass READINI(file,"section", ""), function returns array of all keys
' under "section"
'
' If you pass READINI(file,"",""), function returns array of sections in the
' file
'
'===================

'Usage
' READINI (file, section, item) returns value; otherwise returns ""
'
'
'Example:
' Filename = file.ini
'   [Default]
'   Groups=C:\Warehouse\GROUPS
'
' WScript.Echo ReadINI("file.ini","Default","Groups")
'
' or
'
' ReadVar = ReadINI("file.ini","Default","Groups")
' WScript.Echo ReadVar
'
' To read array of sections
' arrTest = ReadINI("msdfmap.ini","connect CustomerDatabase","Access")
' WScript.Echo arrTest
'
' To read array of keys
' arrTest = ReadINI("msdfmap.ini","connect CustomerDatabase","Access")
' For x=0 to UBound(arrTest)
'     WScript.Echo arrTest(x)
' Next
'
Function ReadINI(file, section, key)
   set FileSysObj = CreateObject("Scripting.FileSystemObject")
   ReadIni=""
   If FileSysObj.FileExists(file) then
       Set ini = FileSysObj.OpenTextFile( file, 1, False)
       ' Return array of sections if section and keys are empty
       if section="" then
           set reSection          =new RegExp
           reSection.Global =True
           reSection.IgnoreCase=True
           'reSection.Pattern ="\[([a-zA-Z0-9 ]*)\]"
           'reSection.Pattern ="\[([a-zA-Z0-9_]*)\]"
           'reSection.Pattern = "\[(.*?)\]"
           'reSection.Pattern = "^\s*\[.+\]\s*"
           reSection.Pattern = "^\s*\[(.+)\]"
           Do While ini.AtEndofStream =False
               line = ini.ReadLine
               if reSection.Test(line) then
                   'tempSection=tempSection & reSection.Replace(line, "$1") & ","
                   Set colMatches = reSection.Execute(line)
                   For Each objMatch in colMatches
                       strSection = objMatch.Value
                       strSection = Mid(strSection, InStr(strSection,"[") + 1) 'trim off everything before the leading [
                       strSection = Left(strSection, Len(strSection) - 1) 'trim off the trailing ]
                       tempSection=tempSection & strSection & ","
                   Next
               end if
           loop
           ini.close
           if tempSection & "" = "" then exit function
           tempSection=left(tempSection, len(tempSection)-1)
           ReadINI=split(tempSection,",")
           set reSection=nothing
           exit function
       End If
       ' Return array of keys if keys are empty
       If key="" then
           set reSection          =new RegExp
           reSection.Global =True
           reSection.IgnoreCase=True
           reSection.Pattern ="^\s*\[\s*" & section & "\s*\]"
           set reSectionAny          =new RegExp
           reSectionAny.Global =False
           reSectionAny.IgnoreCase=True
           reSectionAny.Pattern = "^\s*\[(.+)\]"
           set reEmpty =new RegExp
           reEmpty.global=true
           reEmpty.IgnoreCase=True
           reEmpty.Pattern ="^\s*$"
           Do While ini.AtEndofStream =False
               line = ini.ReadLine
               if reSection.Test(line) then
                   line=ini.ReadLine
                   do while reSectionAny.Test(line) = False and ini.AtEndofStream = False
                       tempKeys = tempKeys & trim(left(line,instr(line,"=")-1)) & ","
                       line=ini.ReadLine
                       Do While reEmpty.Test(Line) and ini.AtEndofStream = False
                           Line=ini.ReadLine
                       Loop
                   loop
                   tempKeys=Left(tempKeys,(len(tempkeys)-1)) ' Remove last comma
                   ReadINI =split(tempKeys,",")
                   exit function
               end if
           loop
       end if
   '===================
   ' READINI Part for file, section, key
       set reSection          =new RegExp
       reSection.Global =False
       reSection.IgnoreCase=True
       'reSection.Pattern ="\s*[\s*" & section & "\s*]"
       reSection.Pattern ="^\s*\[\s*" & section & "\s*\]"
       set reSectionAny          =new RegExp
       reSectionAny.Global =False
       reSectionAny.IgnoreCase=True
       reSectionAny.Pattern = "^\s*\[(.+)\]"
       set reKey           =new RegExp
       reKey.Global     =False
       reKey.IgnoreCase=True
       reKey.Pattern="^\s*" & key & "\s*=\s*"

       Do While ini.AtEndofStream = False
           line = ini.ReadLine
           if reSection.Test(line) = True then
               line=ini.ReadLine
               do while reSectionAny.Test(line) = False and ini.AtEndofStream = False
                   if reKey.Test(line) then
                       ReadINI=trim(mid(line,instr(line,"=")+1))
                       exit do
                   end if
                   line=ini.ReadLine
               Loop
               exit do 'While ini.AtEndofStream = False
           end if
       loop
       ini.Close
       set reSection=nothing
       set reKey =nothing
   end if ' If FileSysObj
End Function
'==================
' WRITEINI ( file, section, item, value )
' file = path and name of ini file
' section = [Section] must be in brackets in the ini file
' item = the variable to read;
' value = the value to assign to the item.
'
Sub WriteIni( file, section, item, value )
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Dim FileSysObj, read_ini, write_ini
   Dim in_section, section_exists, item_exists, wrote, path, line
   Dim reWSection, reItem
   Dim fsoIniTemp, tsoIniOrig, tsoIniTemp, strIniContents
  
   set FileSysObj = CreateObject("Scripting.FileSystemObject")
   in_section = False
   section_exists = False
   item_exists = ( ReadIni( file, section, item ) <> "" )
   wrote = False
   path = Mid( file, 1, InStrRev( file, "\" ) )
   Set read_ini = FileSysObj.OpenTextFile( file, 1, True, TristateFalse )
   Set write_ini = FileSysObj.CreateTextFile( path & "temp_ini.ini", False )
   set reWSection           =new RegExp
   reWSection.Global      =False
   reWSection.IgnoreCase=True
   reWSection.Pattern      ="^\s*\[\s*" & section & "\s*\]"
   set reItem =new RegExp
   reItem.Global =False
   reItem.IgnoreCase=True
   reItem.Pattern ="^\s*" & item & "\s*="
   While read_ini.AtEndOfStream = False
   line = read_ini.ReadLine
       If wrote = False Then
           If reWSection.Test(line) Then
               section_exists = True
               in_section = True
           ElseIf InStr( line, "[" )> 0 Then
               in_section = False
           End If
       End If
       If in_section Then
           If item_exists = False Then
               write_ini.WriteLine line
               write_ini.WriteLine item & "=" & value
               wrote = True
               in_section = False
               'msgbox "Writing " & line
           ElseIf reItem.Test(line) Then
               write_ini.WriteLine item & "=" & value
               wrote = True
               in_section = False
           Else
               write_ini.WriteLine line
           End If
       Else
           write_ini.WriteLine line
       End If
   Wend
   If section_exists = False Then ' section doesn't exist
       section=trim(section)
       item   =trim(item)
       write_ini.WriteLine
       write_ini.WriteLine "[" & section & "]"
       write_ini.WriteLine item & "=" & value
   End If
   read_ini.Close
   write_ini.Close
   'FileSysObj.DeleteFile file
   'FileSysObj.MoveFile path & "temp_ini.ini", file
  
   Set fsoIniTemp = FileSysObj.GetFile(path & "temp_ini.ini")
   If fsoIniTemp.Size > 0 Then
       Set tsoIniOrig = FileSysObj.OpenTextFile(file, ForWriting)
       Set tsoIniTemp = FileSysObj.OpenTextFile(fsoIniTemp.Path, ForReading)
       strIniContents = tsoIniTemp.ReadAll
       tsoIniOrig.Write strIniContents
       tsoIniOrig.Close
       tsoIniTemp.Close
   Else
       'WScript.Echo file & " is empty. Please rectify."
   End If
   fsoIniTemp.Delete
   set reWSection=nothing
   set reItem=nothing
End Sub
%>