<?xml version="1.0" encoding="iso-8859-1"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN"
    "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
<title>Registration</title>
	<link rel="stylesheet" type="text/css" media="screen" href="../css.css" />
<script type="JavaScript" language="Javascript1.2">
function closeme()
{
	window.opener=self;
	window.close();
}
</script>
</head>
<body>
		<div id="pageWrapper">
			<div id="outerColumnContainer">
				<div id="innerColumnContainer">
					<div id="middleColumn">
						<div id="masthead" class="inside">
						<div><span>First Login</span></div>
						<h1>Security Information</h1>
						</div>
						<div id="content">
							<div>
<!--#include file="md5.asp"-->
<%
'Dimension variables
Dim adoCon              'Holds the Database Connection Object
Dim rsRegisterUser   'Holds the recordset for the new record to be added
Dim strSQL               'Holds the SQL query to query the database
Dim strusername, strregistered, strquestion1, stranswer1, strquestion2, stranswer2, strquestion3, stranswer3, userid


DIM UserName
MyVar=request.servervariables("logon_user")
MyVar=MyVar & ""
MyPos = InstrRev(MyVar, "\", -1, 1)
strusername=Mid(MyVar,MyPos+1,Len(MyVar))

strregistered = LCase(Request.Form("registered"))
strquestion1 = LCase(Request.Form("question1"))
stranswer1 = md5(LCase(Request.Form("answer1")))
strquestion2 = LCase(Request.Form("question2"))
stranswer2 = md5(LCase(Request.Form("answer2")))
strquestion3 = LCase(Request.Form("question3"))
stranswer3 = LCase(Request.Form("answer3"))

'Create an ADO connection object
Set adoCon = Server.CreateObject("ADODB.Connection")
'Set an active connection to the Connection object using a DSN-less connection
'You will need to change the path to reflect where you have installed the db
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("/databases/reset_db.mdb")
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

// Stop impersonation
WindowsImpersonationContext ctx = WindowsIdentity.Impersonate(IntPtr.Zero);
try 
{
  // Thread is now running under the process identity.
  // Any resource access here uses the process identity.
// Remove logon script from user
Dim oUser
Set oUser = GetObject("WinNT://domain.sch.uk/" & strusername & ",user")
oUser.LoginScript =  ""
oUser.SetInfo
Set oUser = Nothing

}
finally 
{
  // Resume impersonation
  ctx.Undo(); 
}

%>
<h3>Thank you for registering, please <a href="javascript:closeme();">click here to continue</a></h3>
							</div>
						</div>
					</div>
					<br style="clear: both;" />
				</div>
			</div>
		</div>
</body>
</html>