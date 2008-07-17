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
PageTitle = "Information Update"
PageSubTitle = "Step 2"
PageMenu = False

On Error Resume Next
%>
<!--#include virtual="/config.asp"-->
<!--#include virtual="/md5.asp"-->
<!--#include virtual="/passtype.asp"-->
<%
Dim oDomain
Set oDomain = GetObject("WinNT://" & FQDN)
Dim PageContent(0)

JavaScriptHeader = "function formValidation(form)" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.NewPass))" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(minLength(form.NewPass)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Your password must be at least " & oDomain.MinPasswordLength & " characters long"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.NewPass.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Please enter a new password"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.NewPass.focus();" & vbNewLine
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

'Need to connect to database and verify the answers are correct if not tell them to go back, if correct allow them to change pasword.
'Dimension variables
Dim adoCon         'Holds the Database Connection Object
Dim rsReset   'Holds the recordset for the records in the database
Dim strSQL          'Holds the SQL query to query the database 
Dim strusername 'Holds the username passed from the hidden field.  
Dim strregistered
Dim strquestion1
Dim stranswer1
Dim strquestion2
Dim stranswer2
Dim strquestion3
Dim stranswer3
Dim NewPassword

'strusername = Request.Form("username") 

'Stores username submitted by form into a variable.   
strregistered = Request.Form("registered")
strquestion1 = Request.Form("question1")
stranswer1 = Request.Form("answer1")
strquestion2 = Request.Form("question2")
stranswer2 = Request.Form("answer2")
strquestion3 = Request.Form("question3")
stranswer3 = Request.Form("answer3")

If HashAnswers = True Then
stranswer1 = md5(LCase(stranswer1))
stranswer2 = md5(LCase(stranswer2))
Else
stranswer1 = LCase(stranswer1)
stranswer2 = LCase(stranswer2)
End If
stranswer3 = LCase(stranswer3)

'Stores username submitted by form into a variable.   
MyVar=request.servervariables("logon_user")
MyVar=MyVar & ""
MyPos = InstrRev(MyVar, "\", -1, 1)
strusername=Mid(MyVar,MyPos+1,Len(MyVar))

strusername = LCase(strusername)

'Check for form input
If (Not IsNull(strusername)) AND (Not IsNull(stranswer1)) AND (Not IsNull(stranswer2)) AND (Not IsNull(stranswer3)) Then

'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection")
'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";" 
'Create an ADO recordset object
Set rsReset = Server.CreateObject("ADODB.Recordset")
'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblmain.* FROM tblmain where username='" & strusername & "';"
'Open the recordset with the SQL query 
rsReset.Open strSQL, adoCon
'Loop through the recordset
If Not rsReset.EOF Then
Do While Not rsReset.EOF
	     'Write the code to confirm what was typed
if stranswer1 = rsReset("Answer1") and stranswer2 = rsReset("Answer2") and stranswer3 = rsReset("Answer3") then

	'Delete user from the database
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
	Response.Redirect "/register/register.asp"

else
	PageContent(0) = "One or more answers incorrect press back and try again<br /><a href=""index.asp"">Go Back</a>"
end if
     'Move to the next record in the recordset
     rsReset.MoveNext

Loop

Else
'User not in database
	Response.Redirect "/register/register.asp"
End If
'Reset server objects
rsReset.Close
Set rsReset = Nothing
Set adoCon = Nothing

Else
PageContent(0) = "<h3>Error. Form did not provide expected values.<br /><a href=""index.asp"">Please try again</a><br />If you recieve this message again, please contact an "
	Select Case AdminLinkType
		Case 1 'No link
		PageContent(0) = PageContent(0) & "Administrator"
		Case 2 'Mailto:
		PageContent(0) = PageContent(0) & "<a href=""mailto:" & AdminLinkEmail & """>Administrator</a>"
		Case 3 'Web Link
		PageContent(0) = PageContent(0) & "<a href=""http://" & AdminLinkWeb & """>Administrator</a>"
		Case Else
		PageContent(0) = PageContent(0) & "Administrator"
	End Select
PageContent(0) = PageContent(0) & "</h3>"
End If
%>
<!--#include virtual="/template.asp"-->