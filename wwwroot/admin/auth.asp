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
'-----
MyVar=request.servervariables("logon_user")
MyVar=MyVar & ""
MyPos = InstrRev(MyVar, "\", -1, 1)
CurrentUser=Mid(MyVar,MyPos+1,Len(MyVar))

CurrentUser = LCase(CurrentUser)

Set objLogon = Server.CreateObject("LoginAdmin.ImpersonateUser")
objLogon.Logon ImpersonateUser, ImpersonateUserPass, FQDN
Set UserObj = GetObject("WinNT://" & FQDN & "/" & CurrentUser & ",user")
objLogon.Logoff
Set objLogon = Nothing

For Each GroupObj in UserObj.Groups
    If LCase(GroupObj.Name) = LCase(adminGroup) Then
		AuthAdmin = True
	End If
    If LCase(GroupObj.Name) = LCase(resetGroup) Then
		AuthReset = True
    End If
Next
'-----
End If
%>