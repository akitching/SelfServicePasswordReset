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

On Error Resume Next
%>
<!--#include virtual="/config.asp"-->
<!--#include file="auth.asp"-->
<!--#include virtual="/functions.asp"-->
<!--#include file="menu.asp"-->
<%
JavaScriptHeader = "function formValidation(form){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.answer)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Did the user not give you an answer?"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.verify.answer.focus();" & vbNewLine
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

	Dim strusername, Domain, arrDomain
	strusername = Request.Form("username")
	strusername = LCase(strusername)
	
	Set oUser = GetObject("WinNT://" & FQDN & "/" & strusername & ",user")
	If Err.number <> 0 Then
	PageSubTitle = "...Error..."
		If Err.Number = -2147022675 Then
			PageContent(0) = "<h3>Error: User " & strusername & " does not exist</h3>"
		Else
            If AdminName = "" Then
			    PageContent(0) = "An unexpected error occurred in the execution of this page<br />Please report the following information to an "
                AdminName = "Administrator"
            Else
                PageContent(0) = "An unexpected error occurred in the execution of this page<br />Please report the following information to "
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
			PageContent(0) = PageContent(0) & "<br /><b>Page Error Object</b><br />Error Number: " & Err.Number & "<br />Error Description: " & Err.Description & "<br />Source: " & Err.Source & "<br />LineNumber: " & Err.HelpContext & "</h3>"
		End If
	Else

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
	
	arrDomain = Split(FQDN, ".")
	If Len(arrDomain(0)) > 11 Then
		'Shorten to 11 characters
		NTDomain = UCase(Left(arrDomain(0), 11))
	Else
		NTDomain = UCase(arrDomain(0))
	End If
	
	'Create an ADO connection object
	Set adoCon = Server.CreateObject("ADODB.Connection")
	'Set an active connection to the Connection object using a DSN-less connection
	adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";"

	'Create an ADO recordset object
	Set rsReset = Server.CreateObject("ADODB.Recordset")
    Set rsReset2 = Server.CreateObject("ADODB.Recordset")
	'Initialise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT tblmain.* FROM tblmain where username='" & strusername & "';"
    strSQL2 = "SELECT userdata.* FROM userdata where username='" & strusername & "';"
	'Open the recordset with the SQL query 
	rsReset.Open strSQL, adoCon
    rsReset2.Open strSQL2, adoCon

PageContent(0) = "<form method=""post"" name=""verify"" onsubmit=""return formValidation(this)"" action=""search.asp"">" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""hidden"" name=""username"" value=""" & oUser.Name & """ />" & vbNewLine
PageContent(0) = PageContent(0) & "<table>" & vbNewLine
PageContent(0) = PageContent(0) & "<tr><td>SAMAccountName</td><td> : </td><td>" & NTDomain & "\" & oUser.Name & "</td></tr>" & vbNewLine
PageContent(0) = PageContent(0) & "<tr><td>Display Name</td><td> : </td><td>" & oUser.FullName & "</td></tr>" & vbNewLine
If AuthAdmin Then
    DN = SearchDistinguishedName(oUser.Name)
    PageContent(0) = PageContent(0) & "<tr><td colspan=""3"">" & DN & "</td></tr>" & vbNewLine
    PageContent(0) = PageContent(0) & "<tr><td>Login ~~Script~~</td><td> : </td><td>" & oUser.LoginScript & "</td></tr>" & vbNewLine
    PageContent(0) = PageContent(0) & "<tr><td>Description</td><td> : </td><td>" & oUser.Description & "</td></tr>" & vbNewLine
    PageContent(0) = PageContent(0) & "<tr><td>Home Directory</td><td> : </td><td>" & oUser.HomeDirectory & "</td></tr>" & vbNewLine
    PageContent(0) = PageContent(0) & "<tr><td>Profile Path</td><td> : </td><td>" & oUser.Profile & "</td></tr>" & vbNewLine
    PageContent(0) = PageContent(0) & "<tr><td>Account Locked</td><td> : </td><td>" & oUser.IsAccountLocked & "</td></tr>" & vbNewLine
    PageContent(0) = PageContent(0) & "<tr><td>Account Disabled</td><td> : </td><td>" & oUser.AccountDisabled & "</td></tr>" & vbNewLine
    PageContent(0) = PageContent(0) & "<tr><td>Password Last Set</td><td> : </td><td>"
    intPasswordAge = oUser.PasswordAge
    If intPasswordAge > 0 Then
        intPasswordAge = intPasswordAge * -1 
        dtmChangeDate = DateAdd("s", intPasswordAge, Now)
        PageContent(0) = PageContent(0) & dtmChangeDate
    Else If oUser.PasswordExpired = 1 Then
        PageContent(0) = PageContent(0) & "Password has been expired"
    Else
        PageContent(0) = PageContent(0) & "Password never expires"
    End If
    End If
    PageContent(0) = PageContent(0) & "</td><tr>" & vbNewLine
    If Not rsReset2.EOF Then
        PageContent(0) = PageContent(0) & "<tr><td>Times Reset by Admin</td><td> : </td><td>" & rsReset2("resetcount") & "</td></tr>" & vbNewLine
        PageContent(0) = PageContent(0) & "<tr><td>Last Admin Reset</td><td> : </td><td>" & rsReset2("resetdate") & "</td></tr>" & vbNewLine
    End If
End If

	'Loop through the recordset
	'Do While Not rsReset.EOF
	If Not rsReset.EOF Then
	'Check number of failed password reset attempts
	PageContent(0) = PageContent(0) & "<tr><td>Failed Reset Attempts</td><td> : </td><td>" & rsReset("counter") & "</td></tr>" & vbNewLine
	PageContent(0) = PageContent(0) & "<tr><td>Maximum Reset Attempts</td><td> : </td><td>" & ResetAttempts & "</td></tr>" & vbNewLine
		If rsReset("counter") => ResetAttempts Then
			PageContent(0) = PageContent(0) & "<tr><td>&nbsp;</td><td>&nbsp;</td><td class=""locked"">Locked</td></tr>" & vbNewLine
		Else
			PageContent(0) = PageContent(0) & "<tr><td>&nbsp;</td><td>&nbsp;</td><td class=""unlocked"">Unlocked</td></tr>" & vbNewLine
		End If
		If showSemiSecretAnswer Then
			If AuthAdmin OR resetGroupSemiSecret Then
                If HashAnswers = False Then
                PageContent(0) = PageContent(0) & "<tr><td>Question 1</td><td> : </td><td>" & QuestionArray(rsReset("Question1")) & "</td></tr>" & vbNewLine
				PageContent(0) = PageContent(0) & "<tr><td>Answer</td><td> : </td><td>" & rsReset("Answer1") & "</td></tr>" & vbNewLine
                PageContent(0) = PageContent(0) & "<tr><td>Question 2</td><td> : </td><td>" & QuestionArray(rsReset("Question2")) & "</td></tr>" & vbNewLine
				PageContent(0) = PageContent(0) & "<tr><td>Answer</td><td> : </td><td>" & rsReset("Answer2") & "</td></tr>" & vbNewLine
                PageContent(0) = PageContent(0) & "<tr><td>Quetion 3</td><td> : </td><td>" & QuestionArray(rsReset("Question3")) & "</td></tr>" & vbNewLine
				PageContent(0) = PageContent(0) & "<tr><td>Answer</td><td> : </td><td>" & rsReset("Answer3") & "</td></tr>" & vbNewLine
                Else
				PageContent(0) = PageContent(0) & "<tr><td>Semi-secret question</td><td> : </td><td>" & QuestionArray(rsReset("Question3")) & "</td></tr>" & vbNewLine
				PageContent(0) = PageContent(0) & "<tr><td>Semi-secret answer</td><td> : </td><td>" & rsReset("Answer3") & "</td></tr>" & vbNewLine
                End If
			End If
		Else
			If (Len(Request.Form("answer" & "")) <> 0) Then
				If rsReset("answer3") = LCase(Request.Form("answer")) Then
					'Verified
					PageContent(0) = PageContent(0) & "<tr><td colspan=""2"">&nbsp;</td><td class=""unlocked"">User Verified</td></tr>" & vbNewLine
				Else
					'Not Verified
					PageContent(0) = PageContent(0) & "<tr><td colspan=""2"">&nbsp;</td><td class=""locked"">User Not Verified</td></tr>" & vbNewLine
					PageContent(0) = PageContent(0) & "<tr><td>Semi-secret question</td><td> : </td><td>" & QuestionArray(rsReset("Question3")) & "</td></tr>" & vbNewLine
					PageContent(0) = PageContent(0) & "<tr><td><label for=""answer"">Semi-secret answer</label></td><td> : </td><td><input type=""text""  name=""answer"" id=""answer"" style=""width: 175px;""  /></td></tr>" & vbNewLine
					PageContent(0) = PageContent(0) & "<tr><td colspan=""2"">&nbsp;</td><td><input type=""submit"" value=""Verify"" /></td></tr>" & vbNewLine
				End If
			Else
				PageContent(0) = PageContent(0) & "<tr><td>Semi-secret question</td><td> : </td><td>" & QuestionArray(rsReset("Question3")) & "</td></tr>" & vbNewLine
				PageContent(0) = PageContent(0) & "<tr><td><label for=""answer"">Semi-secret answer</label></td><td> : </td><td><input type=""text""  name=""answer"" id=""answer"" style=""width: 175px;""  /></td></tr>" & vbNewLine
				PageContent(0) = PageContent(0) & "<tr><td colspan=""2"">&nbsp;</td><td><input type=""submit"" value=""Verify"" /></td></tr>" & vbNewLine
			End If
		End If
	Else
	PageContent(0) = PageContent(0) & "<tr><td colspan=""3"">User not registered in database.</td></tr>" & vbNewLine
	End If
	
PageContent(0) = PageContent(0) & "</table></form>" & vbNewLine

PageContent(0) = PageContent(0) & "<form method=""post"" name=""reset"" action=""adminreset.asp"">" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""hidden"" name=""username"" value=""" & oUser.Name & """ />" & vbNewLine
PageContent(0) = PageContent(0) & "<input type=""submit"" name=""resetPassword"" value=""Reset Password"" />&nbsp;&nbsp;<input type=""submit"" name=""resetCounter"" value=""Reset Counter"" />&nbsp;&nbsp;<input type=""submit"" name=""removeUserAnswers"" value=""Remove User's Answers"" />&nbsp;" & vbNewLine
PageContent(0) = PageContent(0) & "</form>" & vbNewLine

	If AuthAdmin Then
	PageContent(1) = "<br />User is a member of:" & "<br />"
	PageContent(1) = PageContent(1) & "<table>"
	For Each GroupObj in oUser.Groups
		PageContent(1) = PageContent(1) & "<tr><td>" & GroupObj.Name & "</td><td>(" & GroupObj.Description & ")</td></tr>"
	Next
	PageContent(1) = PageContent(1) & "</table>"
	End If
End If
End If
%>
<!--#include virtual="/template.asp"-->