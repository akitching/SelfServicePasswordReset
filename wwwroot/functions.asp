<%
'    Copyright 2006 Ben Norcutt and Alex Kitching
'
'    This file is part of Self Service Password Reset.
'
'    Self Service Password Reset is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the
'    License, or (at your option) any later version.
'
'    Self Service Password Reset is distributed in the hope that it will
'    be useful, but WITHOUT ANY WARRANTY; without even the implied warranty
'    of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Self Service Password Reset; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

IF Not INCLUDED Then
%>
<!--#include virtual="/template.asp"-->
<%
Else

Public Function SearchDistinguishedName(ByVal vSAN)
    ' Function:     SearchDistinguishedName
    ' Description:  Searches the DistinguishedName for a given SamAccountName
    ' Parameters:   ByVal vSAN - The SamAccountName to search
    ' Returns:      The DistinguishedName Name
    Dim oRootDSE, oConnection, oCommand, oRecordSet

    Set oRootDSE = GetObject("LDAP://rootDSE")
    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Open "Provider=ADsDSOObject;"
    Set oCommand = CreateObject("ADODB.Command")
    oCommand.ActiveConnection = oConnection
    oCommand.CommandText = "<LDAP://" & oRootDSE.get("defaultNamingContext") & _
        ">;(&(objectCategory=User)(samAccountName=" & vSAN & "));distinguishedName;subtree"
    Set oRecordSet = oCommand.Execute
    On Error Resume Next
    SearchDistinguishedName = oRecordSet.Fields("DistinguishedName")
    On Error GoTo 0
    oConnection.Close
    Set oRecordSet = Nothing
    Set oCommand = Nothing
    Set oConnection = Nothing
    Set oRootDSE = Nothing
End Function

End If
%>