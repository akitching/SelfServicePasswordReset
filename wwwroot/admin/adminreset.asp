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
Dim objLogon

On Error Resume Next
%>
<!--#include virtual="/config.asp"-->
<!--#include virtual="/md5.asp"-->
<!--#include virtual="/functions.asp"-->
<!--#include file="auth.asp"-->
<!--#include virtual="/passtype.asp"-->
<!--#include file="menu.asp"-->
<%
Dim oDomain
Set oDomain = GetObject("WinNT://" & FQDN)

JavaScriptHeader = "function formValidation(form)" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.NewPass))" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(minLength(form.NewPass)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Your password must be at least " & oDomain.MinPasswordLength & " characters long"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.reset.NewPass.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Please enter a new password"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.reset.NewPass.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "function notEmpty(elem)" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(elem.value.length == 0){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "function minLength(elem)" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(elem.value.length >= " & oDomain.MinPasswordLength & "){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine

	Dim strusername, Domain, arrDomain
	strusername = Request.Form("username")
	strusername = LCase(strusername)
	
    Set objLogon = Server.CreateObject("LoginAdmin.ImpersonateUser")
    objLogon.Logon ImpersonateUser, ImpersonateUserPass, FQDN
	Set oUser = GetObject("WinNT://" & FQDN & "/" & strusername & ",user")
    objLogon.Logoff
    Set objLogon = Nothing

	For Each GroupObj in oUser.Groups
		If GroupObj.Name = usersGroup Then AuthResetable = True End If
	Next

If (Not AuthAdmin) AND ((Not AuthReset) OR (Not AuthResetable)) Then
	PageSubTitle = "Authorised Access Only"
	If (Not AuthAdmin) AND (Not AuthReset) Then
		PageContent(0) = "<h3>You are not authorised to access this service.<br /><a href=""http://" & homepage & """>Leave</a>.</h3>"
	Else If Not AuthResetable Then
		PageContent(0) = "<h3>You are not authorised to access this user.<br /><a href=""index.asp"">Return to Administration</a>.</h3>"
	End If
	End If
Else
	If AuthAdmin Then
		PageSubTitle = "Full Access Granted"
	Else If AuthReset Then
		PageSubTitle = "Reset Access Granted"
	End If
	End If
	
	If (Request.Form("resetCounter") = "Reset Counter") AND (Not IsNull(Request.Form("username"))) Then
		
		'Create an ADO connection object
		Set adoCon = Server.CreateObject("ADODB.Connection")
		'Set an active connection to the Connection object using a DSN-less connection
		adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";" 
		'Create an ADO recordset object
		'Reset failed reset counter
		'Create an ADO recordset object
		Set rsRegisterUser = Server.CreateObject("ADODB.Recordset")
		'Initialise the strSQL variable with an SQL statement to query the database
		strSQL = "UPDATE tblmain SET counter='0' WHERE username='" & strusername & "';"
		'rsRegisterUser.Open strSQL, adoCon
		adoCon.Execute strSQL

		'rsRegisterUser.Close
		'Set rsRegisterUser = Nothing
		Set adoCon = Nothing

		'End Password Reset Code
		PageContent(0) = "Reset Counter for " & Request.Form("username") & " has been reset<br /><a href=""javascript: history.go(-1)"">Back</a><br /><a href=""index.asp"">Return to Administration</a>"
	
	Else If (Request.Form("resetPassword") = "Reset Password") AND (Not IsNull(Request.Form("username"))) Then
		Dim usr, oUser, oFlags, NewPassword
		
		'Check password reset type and select a value
		Select Case PwdType
			Case 1
			NewPassword = DefPassword
			ResetPass()
			Case 2
			NewPassword = PWList(PasswordList, strusername)
			ResetPass()
			Case 3
			NewPassword = RandomPW(RandomPassword)
			ResetPass()
			Case 4
			NewPassword = RandomPWList(RandomPasswordList)
			ResetPass()
			Case 5
			If (Not ISEmpty(Request.Form("NewPass"))) Then
				NewPassword = Request.Form("NewPass")
				ResetPass()
			Else
				PWPromptAdmin()
			End If
			Case Else
                If AdminName = "" Then
                    PageContent(0) = "Error. Password type not defined.<br />Please contact an "
                    AdminName = "Administrator"
                Else
                    PageContent(0) = "Error. Password type not defined.<br />Please contact "
                End If
				Select Case AdminLinkType
					Case 1 'No link
					PageContent(0) = PageContent(0) & AdminName
					Case 2 'Mailto:
					PageContent(0) = PageContent(0) & "<a href=""mailto:" & AdminLinkEmail & """>" & AdminName & "</a>"
					Case 3 'Web Link
					PageContent(0) = PageContent(0) & "<a href=""http://" & AdminLinkWeb & """>" & AdminName & "</a>"
					Case Else
					PageContent(0) = PageContent(0) & AdminName
				End Select
					PageContent(0) = PageContent(0) & " and report this error.</h3>"
		End Select

	Else If (Request.Form("removeUserAnswers") = "Remove User's Answers") AND (Not IsNull(Request.Form("username"))) Then
		'Remove user from database

		'Create an ADO connection object
		Set adoCon = Server.CreateObject("ADODB.Connection")
		'Set an active connection to the Connection object using a DSN-less connection
		adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";" 
		'Create an ADO recordset object
		'Reset failed reset counter
		'Create an ADO recordset object
		Set rsRegisterUser = Server.CreateObject("ADODB.Recordset")
		'Initialise the strSQL variable with an SQL statement to query the database
		strSQL = "DELETE FROM tblmain WHERE username='" & strusername & "';"
		'rsRegisterUser.Open strSQL, adoCon
		adoCon.Execute strSQL

		'rsRegisterUser.Close
		'Set rsRegisterUser = Nothing
		Set adoCon = Nothing

		'End Password Reset Code
		PageContent(0) = "User " & Request.Form("username") & " has been removed from the database<br /><a href=""javascript: history.go(-1)"">Back</a><br /><a href=""index.asp"">Return to Administration</a>"
				
	Else
		PageContent(0) = "Invalid entry to this page"
	End If
	End If
	End If
	
End If

Function ResetPass
Dim objUser, objFlags, dso, allowedAttributesEffective

		'Begin Password Reset Code
		'Connect to Active Directory, reset and expire password
		Set objLogon = Server.CreateObject("LoginAdmin.ImpersonateUser")
		objLogon.Logon ImpersonateUser, ImpersonateUserPass, FQDN
		Set objUser = GetObject("WinNT://" & FQDN & "/" & strusername & ",user")
		If Err.number <> 0 Then
			PageContent(0) = "Error: Unable to bind to container"
		Else
			PageContent(0) = "Resetting password for <b>" & strusername & "</b> ...<br />" & vbNewLine
			On Error Resume Next
			objUser.GetInfo
			objUser.SetPassword (NewPassword)
			If (Err.number <> 0) Then
				If AdminName = "" Then
					PageContent(0) = PageContent(0) & "An unexpected error occurred in the execution of this page<br />Please report the following information to an "
                    AdminName = "Administrator"
                Else
                    PageContent(0) = PageContent(0) & "An unexpected error occurred in the execution of this page<br />Please report the following information to "
                End If
					Select Case AdminLinkType
						Case 1 'No link
						PageContent(0) = PageContent(0) & AdminName
						Case 2 'Mailto:
						PageContent(0) = PageContent(0) & "<a href=""mailto:" & AdminLinkEmail & """>" & AdminName & "</a>"
						Case 3 'Web Link
						PageContent(0) = PageContent(0) & "<a href=""http://" & AdminLinkWeb & """>" & AdminName & "</a>"
						Case Else
						PageContent(0) = PageContent(0) & AdminName
					End Select
                    allowedAttributesEffective = objUser.GetInfoEx (allowedAttributesEffective)
					PageContent(0) = PageContent(0) & "<br /><b>Page Error Object</b><br />Error Number: " & Err.Number & "<br />Error Description: " & Err.Description & "<br />Source: " & Err.Source & "<br />LineNumber: " & Err.HelpContext & "<br />Allowed Attributes: " & allowedAttributesEffective & "</h3>"
			Else
            objUser.Put "PasswordExpired", CLng(1)
		    objUser.SetInfo
			PageContent(0) = "Resetting password for <b>" & strusername & "</b> ...<br />" & vbNewLine
			PageContent(0) = PageContent(0) & "Password has been reset to <b>" & NewPassword & "</b><br /><a href=""javascript: history.go(-1)"">Back</a><br /><a href=""index.asp"">Return to Administration</a>"
			End If
            On Error GoTo 0
		End If

		objLogon.Logoff
		Set objLogon = Nothing

		Set objUser = Nothing

		On Error Resume Next
		'Create an ADO connection object
		Set adoCon = Server.CreateObject("ADODB.Connection")
		'Set an active connection to the Connection object using a DSN-less connection
		adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";" 
		'Create an ADO recordset object
		'Reset failed reset counter
		'Create an ADO recordset object
		'Set rsRegisterUser = Server.CreateObject("ADODB.Recordset")
		'Initialise the strSQL variable with an SQL statement to query the database
		Set rsReset = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT tblmain.* FROM tblmain WHERE username='" & strusername & "';"
		rsReset.Open strSQL, adoCon
		If Not rsReset.EOF Then
		strSQL = "UPDATE tblmain SET counter='0' WHERE username='" & strusername & "';"
		Else
			strSQL = "INSERT INTO tblmain (username, counter) VALUES ('" & strusername & "', '0')"
		End If
		Set rsReset = Nothing
		'rsRegisterUser.Open strSQL, adoCon
'		adoCon.Execute strSQL
		strSQL = "INSERT INTO resetlog (username, resetdate, resetby) VALUES ('" & strusername & "', '" & Now & "', '" & CurrentUser & "')"
		adoCon.Execute strSQL
        
        Set rsReset = Server.CreateObject("ADODB.Recordset")        
        strSQL = "SELECT userdata.* FROM userdata WHERE username='" & strusername & "';"
        
        rsReset.Open strSQL, adoCon
    	If Not rsReset.EOF Then
            resetCount = rsReset("resetcount") + 1
            strSQL = "UPDATE userdata SET resetcount='" & resetCount & "', resetdate='" & Now & "', resetby ='" & CurrentUser & "' WHERE username='" & strusername & "';"
	Else
            strSQL = "INSERT INTO userdata (username, resetcount, resetdate, resetby) VALUES ('" & strusername & "', '1', '" & Now & "', '" & CurrentUser & "')"
	End If
        
        adoCon.Execute strSQL
        

		Set rsReset = Nothing
		Set adoCon = Nothing

		'End Password Reset Code
End Function
%>
<!--#include virtual="/template.asp"-->