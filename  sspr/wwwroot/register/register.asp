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
PageTitle = "Registration"
PageSubTitle = "Step 1"
PageMenu = False
Dim PageContent(1)
%>
<!--#include virtual="/config.asp"-->
<%
JavaScriptHeader = "function formValidation(form)" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "{" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.answer1)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.answer2)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(notEmpty(form.answer3)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(questionSelected(form.question1)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(questionSelected(form.question2)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(questionSelected(form.question3)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(uniqueQuestions(form)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(uniqueAnswers(form)){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""You have used the same answer for 2 or more questions. Your answers need to be unique"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Questions selected are not unique"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Question 3 not selected"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.question3.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Question 2 not selected"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.question2.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "alert(""Question 1 not selected"");" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "document.form.question1.focus();" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
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
JavaScriptHeader = JavaScriptHeader & "function questionSelected(elem){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(elem.value == ""a""){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "function uniqueQuestions(elem){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(elem.question1.value == elem.question2.value || elem.question1.value == elem.question3.value || elem.question2.value == elem.question3.value){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "function uniqueAnswers(elem){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "if(elem.answer1.value == elem.answer2.value || elem.answer1.value == elem.answer3.value || elem.answer2.value == elem.answer3.value){" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return false;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "} else {" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "return true;" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
JavaScriptHeader = JavaScriptHeader & "}" & vbNewLine
      
'Check if user already exists in database
'Dimension variables
Dim adoCon         'Holds the Database Connection Object
Dim rsUserExists   'Holds the recordset for the records in the database
Dim strSQL         'Holds the SQL query to query the database 
Dim strusername 'Holds the current username  
'strusername = Request.Form("username") 

'Stores username submitted by form into a variable.   
MyVar=request.servervariables("logon_user")
MyVar=MyVar & ""
MyPos = InstrRev(MyVar, "\", -1, 1)
strusername=Mid(MyVar,MyPos+1,Len(MyVar))

strusername = LCase(strusername)

'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection")
'Set an active connection to the Connection object using a DSN-less connection
'You will need to change the path to reflect where you have installed the db
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";"

'Check for strusername in database, if present, issue error
Set rsUserExists = Server.CreateObject("ADODB.Recordset")
'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblmain.* FROM tblmain where username='" & strusername & "';"
'Open the recordset with the SQL query 
rsUserExists.Open strSQL, adoCon
'Check if already in database, if so redirect
If NOT rsUserExists.EOF then
Response.Redirect "http://" & HomePage & "/"
End If

'Reset server objects
rsUserExists.Close
Set rsUserExists = Nothing

i = 0
For Each x in QuestionArray
options = options & "<option value=" & i & ">" & x & "</option>"
i = i+1
Next

PageContent(0) = "Please answer 3 of the questions listed below.<br />Try to give answers that cannot easily be guessed, yet that you will still be able to remember in a few months as you will need to know the answers to these questions should you ever forget your password.<br />Please take care to check your spelling is correct.<br />Protect these answers as you would your password, but do not use your password as one of the answers!" & vbNewLine

PageContent(1) = "<!-- Begin form code -->" & vbNewLine
PageContent(1) = PageContent(1) & "<form name=""form"" method=""post"" onsubmit=""return formValidation(this)"" action=""register_user.asp"">" & vbNewLine
PageContent(1) = PageContent(1) & "<input type=""hidden"" name=""registered"" value=""True"" />" & vbNewLine
PageContent(1) = PageContent(1) & "<table>" & vbNewLine
PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c1"">Username:</td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c2"">" & strusername & "</td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c1"">Question 1:</td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c2""><select name=""question1"">" & vbNewLine
PageContent(1) = PageContent(1) & "		<option value=""a"">-- Please Select a Question --</option>" & vbNewLine
PageContent(1) = PageContent(1) & options & vbNewLine
PageContent(1) = PageContent(1) & "	</select></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c1""><label for=""answer1"">Answer 1:</label></td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c2""><input type=""text"" name=""answer1"" id=""answer1"" /></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine

PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c1"">Question 2:</td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c2""><select name=""question2"">" & vbNewLine
PageContent(1) = PageContent(1) & "		<option value=""a"">-- Please Select a Question --</option>" & vbNewLine
PageContent(1) = PageContent(1) & options & vbNewLine
PageContent(1) = PageContent(1) & "	</select></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c1""><label for=""answer2"">Answer 2:</label></td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c2""><input type=""text"" name=""answer2"" id=""answer2"" /></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine

If HashAnswers = True Then

PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c3"" colspan=""2"">Semi-Private Question:<br />" & vbNewLine
PageContent(1) = PageContent(1) & "When you call the help desk, you may be asked to disclose this answer to verify your identity.</td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine

PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c1"">Question 3:</td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c2""><select name=""question3"">" & vbNewLine
PageContent(1) = PageContent(1) & "		<option value=""a"">-- Please Select a Question --</option>" & vbNewLine
PageContent(1) = PageContent(1) & options & vbNewLine
PageContent(1) = PageContent(1) & "	</select></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c1""><label for=""answer3"">Answer 3:</label></td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c2""><input type=""text"" name=""answer3"" id=""answer3"" /></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine

PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c3"" colspan=""2""><input type=""submit"" value=""Submit"" /></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
PageContent(1) = PageContent(1) & "</table>" & vbNewLine
PageContent(1) = PageContent(1) & "<br style=""clear: both;"" />" & vbNewLine
PageContent(1) = PageContent(1) & "</form>" & vbNewLine
PageContent(1) = PageContent(1) & "<!-- End form code -->" & vbNewLine

Else

PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c1"">Question 3:</td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c2""><select name=""question3"">" & vbNewLine
PageContent(1) = PageContent(1) & "		<option value=""a"">-- Please Select a Question --</option>" & vbNewLine
PageContent(1) = PageContent(1) & options & vbNewLine
PageContent(1) = PageContent(1) & "	</select></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c1""><label for=""answer3"">Answer 3:</label></td>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r1 c2""><input type=""text"" name=""answer3"" id=""answer3"" /></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine

PageContent(1) = PageContent(1) & "<tr>" & vbNewLine
PageContent(1) = PageContent(1) & "	<td class=""r2 c3"" colspan=""2""><input type=""submit"" value=""Submit"" /></td>" & vbNewLine
PageContent(1) = PageContent(1) & "</tr>" & vbNewLine
PageContent(1) = PageContent(1) & "</table>" & vbNewLine
PageContent(1) = PageContent(1) & "<br style=""clear: both;"" />" & vbNewLine
PageContent(1) = PageContent(1) & "</form>" & vbNewLine
PageContent(1) = PageContent(1) & "<!-- End form code -->" & vbNewLine

End If

Set adoCon = Nothing
%>
<!--#include virtual="/template.asp"-->