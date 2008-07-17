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

Dim INCLUDED
INCLUDED = True
PageTitle = "Administration"
PageMenu = True
Dim PageContent(1)
%>
<!--#include virtual="/config.asp"-->
<!--#include file="auth.asp"-->
<!--#include file="menu.asp"-->
<%
JavaScriptHeader = "function formValidation(form){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.username)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Who are you searching for?"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.search.username.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "function notEmpty(elem){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(elem.value.length == 0){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine

If (Not AuthAdmin) AND (Not AuthReset) Then
	PageSubTitle = "Authorised Access Only"
	PageContent(0) = "<h3>You are not authorised to access this service.<br /><a href=""http://" & homepage & """>Leave</a>.</h3>"
Else
	If AuthAdmin Then
		PageSubTitle = "Full Access Granted"
	Else If AuthReset Then
		PageSubTitle = "Reset Access Granted"
	End If
	End If
		PageContent(0) = "Enter the logon name of the user whose password you wish to reset<br />This service can only find exact usernames"
		
		PageContent(1) = "<form method=""post"" id=""search"" onsubmit=""return formValidation(this)"" action=""search.asp"">" & vbNewLine
		PageContent(1) = PageContent(1) & "<table>" & vbNewLine
		PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
		PageContent(1) = PageContent(1) & "<td><input type=""text"" name=""username"" style=""width: 175px;""  /></td>" & vbNewLine
		PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
		PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
		PageContent(1) = PageContent(1) & "<td>" & vbNewLine
		PageContent(1) = PageContent(1) & "<select name=""search_type"" style=""width: 175px;"">" & vbNewLine
		PageContent(1) = PageContent(1) & "<option value=""user"" selected=""selected"">User</option>" & vbNewLine
		PageContent(1) = PageContent(1) & "</select>" & vbNewLine
		PageContent(1) = PageContent(1) & "</td>" & vbNewLine
		PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
		PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
		PageContent(1) = PageContent(1) & "<td colspan=""3""><input type=""submit"" value=""Search"" /></td>" & vbNewLine
		PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
		PageContent(1) = PageContent(1) & "</table>" & vbNewLine
		PageContent(1) = PageContent(1) & "</form>" & vbNewLine
End If
%>
<!--#include virtual="/template.asp"-->