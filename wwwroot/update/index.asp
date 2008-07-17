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
PageSubTitle = "Step 1"
Dim PageContent(1)
%>
<!--#include virtual="/config.asp"-->
<%
JavaScriptHeader = "function formValidation(form)" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.answer1)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.answer2)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.answer3)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Answer 3 cannot be empty"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.answer3.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Answer 2 cannot be empty"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.answer2.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Answer 1 cannot be empty"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.answer1.focus();" & vbNewLine
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
'Dimension variables
Dim adoCon         'Holds the Database Connection Object
Dim rsReset   'Holds the recordset for the records in the database
Dim strSQL          'Holds the SQL query to query the database 
dim strusername

'Stores username submitted by form into a variable.   
MyVar=request.servervariables("logon_user")
MyVar=MyVar & ""
MyPos = InstrRev(MyVar, "\", -1, 1)
strusername=Mid(MyVar,MyPos+1,Len(MyVar))

strusername = LCase(strusername)

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
'Do While Not rsReset.EOF
If Not rsReset.EOF Then
'Check number of failed password reset attempts
If rsReset("counter") => ResetAttempts Then
'Dim PageContent(0)
PageContent(0) = "<h2>Maximum number of failed password reset attempts exceeded<br />This service must be unlocked before you can update your details<br />Please contact your system "
	Select Case AdminLinkType
		Case 1 'No link
		PageContent(0) = PageContent(0) & "Administrator"
		Case 2 'Mailto:
		PageContent(0) = PageContent(0) & "<a href=""mailto:" & AdminLinkEmail & """>Administrator</a>"
		Case 3 'Web Link
		PageContent(0) = PageContent(0) & "<a href=""mailto:" & AdminLinkWeb & """>Administrator</a>"
		Case Else
		PageContent(0) = PageContent(0) & "Administrator"
	End Select
	PageContent(0) = PageContent(0) & "</h2>"
Else
'Dim PageContent(1)
PageContent(0) = "Please answer the questions below making sure that the answers you give are the same as the ones you gave when first logging on to your account."

PageContent(1) = "<form name=""form"" method=""post"" onSubmit=""return formValidation(this)"" action=""update_details.asp"">"
PageContent(1) = PageContent(1) & "<input type=""hidden"" name=""username"" maxlength=""50"" value=""" & strusername & """ /><table>"
PageContent(1) = PageContent(1) & "<tr><td class=""r1 c1"">Question 1:</td><td class=""r1 c2"">" & QuestionArray(rsReset("Question1")) & "</td></tr>"
PageContent(1) = PageContent(1) & "<tr><td class=""r2 c1""><label for=""answer1"">Answer 1:</label></td><td class=""r2 c2""><input type=""password"" name=""answer1"" id=""answer1"" maxlength=""50"" /></td></tr>"
PageContent(1) = PageContent(1) & "<tr><td class=""r1 c1"">Question 2:</td><td class=""r1 c2"">" & QuestionArray(rsReset("Question2")) & "</td></tr>"
PageContent(1) = PageContent(1) & "<tr><td class=""r2 c1""><label for=""answer2"">Answer 2:</label></td><td class=""r2 c2""><input type=""password"" name=""answer2"" id=""answer2"" maxlength=""50"" /></td></tr>"
PageContent(1) = PageContent(1) & "<tr><td class=""r1 c1"">Question 3:</td><td class=""r1 c2"">" & QuestionArray(rsReset("Question3")) & "</td></tr>"
PageContent(1) = PageContent(1) & "<tr><td class=""r2 c1""><label for=""answer3"">Answer 3:</label></td><td class=""r2 c2""><input type=""password"" name=""answer3"" id=""answer3"" maxlength=""50"" /></td></tr>"
PageContent(1) = PageContent(1) & "<tr><td class=""r1 c3"" colspan=""2""><input type=""submit"" value=""Submit"" /></td></tr></table></form>"

End If ' If rsReset("counter") => ResetAttempts Then
     'Move to the next record in the recordset
     rsReset.MoveNext 
Else ' If Not rsReset.EOF Then
	Response.Redirect "/register/register.asp"
End If ' If Not rsReset.EOF Then
'Reset server objects
rsReset.Close
Set rsReset = Nothing
Set adoCon = Nothing
%>
<!--#include virtual="/template.asp"-->