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
PageSubTitle = "Step 2"
Dim PageContent(0)
%>
<!--#include virtual="/config.asp"-->
<!--#include virtual="/md5.asp"-->
<%
'Dimension variables
Dim adoCon              'Holds the Database Connection Object
Dim rsRegisterUser   'Holds the recordset for the new record to be added
Dim rsUserExists
Dim strSQL               'Holds the SQL query to query the database
Dim strSQL2           'Holds SQL for username query
Dim strusername, strregistered, strquestion1, stranswer1, strquestion2, stranswer2, strquestion3, stranswer3, userid

DIM UserName
MyVar=request.servervariables("logon_user")
MyVar=MyVar & ""
MyPos = InstrRev(MyVar, "\", -1, 1)
strusername=Mid(MyVar,MyPos+1,Len(MyVar))

strusername = LCase(strusername)

strregistered = LCase(Request.Form("registered"))
strquestion1 = LCase(Request.Form("question1"))
stranswer1 = LCase(Request.Form("answer1"))
strquestion2 = LCase(Request.Form("question2"))
stranswer2 = LCase(Request.Form("answer2"))
strquestion3 = LCase(Request.Form("question3"))
stranswer3 = LCase(Request.Form("answer3"))

'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection")
'Set an active connection to the Connection object using a DSN-less connection
'You will need to change the path to reflect where you have installed the db
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DataBase & ";uid=" & DBUID & ";pwd=" & DBPWD & ";"

'Check for strusername in database, if present, issue error
Set rsUserExists = Server.CreateObject("ADODB.Recordset")
'Initialise the strSQL variable with an SQL statement to query the database
strSQL2 = "SELECT tblmain.* FROM tblmain where username='" & strusername & "';"
'Open the recordset with the SQL query 
rsUserExists.Open strSQL2, adoCon
'Check if already in database, if so error
If NOT rsUserExists.EOF Then
	If AdminName = "" Then
        PageContent(0) = "<h3>Error. User already exists in database<br />You can only register once<br /><a href=""register.asp"">Click here to continue</a><br />If you did not register this account please contact an "
        AdminName= "Administrator"
    Else
        PageContent(0) = "<h3>Error. User already exists in database<br />You can only register once<br /><a href=""register.asp"">Click here to continue</a><br />If you did not register this account please contact "
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
	PageContent(0) = PageContent(0) & "</h3>"
Else

'Check for required form data
If (Not IsNull(strusername)) AND (Not IsNull(strquestion1))  AND IsNumeric(strquestion1) AND (Not IsNull(stranswer1)) AND (Not IsNull(strquestion2)) AND IsNumeric(strquestion2) AND (Not IsNull(stranswer2)) AND (Not IsNull(strquestion3)) AND IsNumeric(strquestion3) AND (Not IsNull(stranswer3)) then


If HashAnswers = True Then

'MD5 Hash secret answers
stranswer1 = md5(stranswer1)
stranswer2 = md5(stranswer2)

End If

'Create an ADO recordset object
Set rsRegisterUser = Server.CreateObject("ADODB.Recordset")
'Initialise the strSQL variable with an SQL statement to query the database
'strSQL = "SELECT tblmain.username, tblmain.registered FROM tblmain;"
strSQL = "SELECT tblmain.* FROM tblmain;"
'Set the cursor type we are using so we can navigate through the recordset
rsRegisterUser.CursorType = 2
'Set the lock type so that the record is locked by ADO when it is updated
rsRegisterUser.LockType = 3
'Open the recordset with the SQL query 
rsRegisterUser.Open strSQL, adoCon
'Tell the recordset we are adding a new record to it
rsRegisterUser.AddNew
'Add a new record to the recordset
rsRegisterUser.Fields("username") = strusername
rsRegisterUser.Fields("registered") = strregistered
rsRegisterUser.Fields("question1") = strquestion1
rsRegisterUser.Fields("answer1") = stranswer1
rsRegisterUser.Fields("question2") = strquestion2
rsRegisterUser.Fields("answer2") = stranswer2
rsRegisterUser.Fields("question3") = strquestion3
rsRegisterUser.Fields("answer3") = stranswer3
'Write the updated recordset to the database
rsRegisterUser.Update
'Reset server objects
rsRegisterUser.Close
Set rsRegisterUser = Nothing
Set adoCon = Nothing


PageContent(0) = "<h3>Thank you for registering, please <a href=""register.asp"">click here to continue</a></h3>"
else
IF AdminName = "" Then
    PageContent(0) = "<h3>Error. Form did not provide expected values.<br /><a href=""register.asp"">Please try again</a><br />If you recieve this message again, please contact an "
    AdminName = "Administrator"
ELSE
    PageContent(0) = "<h3>Error. Form did not provide expected values.<br /><a href=""register.asp"">Please try again</a><br />If you recieve this message again, please contact "
END IF
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
PageContent(0) = PageContent(0) & "</h3>"
end if
End If

'Reset server objects
rsUserExists.Close
Set rsUserExists = Nothing
%>
<!--#include virtual="/template.asp"-->