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

'Global configuration file for Self Service Password Reset
'Do not change the following lines
IF Not INCLUDED Then
%>
<!--#include virtual="/template.asp"-->
<%
Else
Dim FQDN, Administrator, HomePage, ResetAttempts, PwdType, DefPassword, PasswordList, RandomPassword, RandomPasswordList, DataBase, DBUID, DBPWD, adminUser, adminPass, HashAnswers, ImpersonateUser, ImpersonateUserPass
'-------------------------------------------------------------------------------
'Change the following values to suit those used in your domain
'-----------------
'FQDN (Fully Qualified Domain Name) of your organisation Eg: office.example.org
FQDN = "office.example.org"
'-----------------
'Location of credentials file for user impersonation
' !IMPORTANT! This file should NOT be in a web accessible directory. You should also restrict access to this file.
CredFile = "c:\Inetpub\cred.ini"
'-----------------
'How should references to the Administrator be linked?
'1) No link
'2) Mailto: link
'3) Web address (Ie: A contact form)
AdminLinkType = 1
'Modify the variable which corresponds to the AdminLinkType chosen above
AdminLinkEmail = "admin@example.org"			'E-Mail address of Password Reset Administrator
AdminLinkWeb = "office.example.org/contact/"	'Web address to link to (Not including http://)
'How should references to the Administrator be displayed? (Leave blank for Administrator)
AdminName = ""
'-----------------
'User HomePage. Do not include http:// Eg: www.google.com
HomePage = "www.google.com"
'-----------------
'Maximum number of reset attempts a user can fail before needing an Administrator to unlock the system.
'Entering a value of 0 will break the script
ResetAttempts = 3
'-----------------
'Is the user forced to reregister after using SSPR to reset their password?
removeOnReset = False
'-----------------
'Should user's answers be help as plain text or should the first two be hashed (for authentication reasons third answer is never hashed)
HashAnswers = False
'-----------------
'Reset password can take 5 forms:
'1) a static password used for all users
'2) a list of user/password pairs
'3) a random alphanumeric srting
'4) a random password chosen from a list
'5) a password chosen by the user
'Set PwdType to the number of the method desired
PwdType = 5
'Modify the variable which corresponds to the PwdType chosen above
DefPassword = "paradise"	'Static password for all users
PasswordList = "C:\Inetpub\userpasslist.csv"		'Path to csv file containing user/password pairs
RandomPassword = 10		'Length of random password. Set to 0 for variable length using the min/max values below
RandomPasswordMin = 6	'Minimum length of random password
RandomPasswordMax = 20	'Mazimum length of random password
RandomPasswordList = "C:\Inetpub\passlist.txt"	'Path to file containing password list, 1 password per line
'-----------------
'Location of database Eg: C:\Inetpub\wwwroot\databases\reset_db.mdb
DataBase = "C:\Inetpub\databases\reset_db.mdb"
'-----------------
'Database Access Credentials.
'Only set these if you have protected the database. The standard database does not need these to be modified
'UserID
DBUID = ""
'Password
DBPWD = ""
'-----------------
'Active Directory Security Groups with rights to administrate SSPR
'adminGroup can reset any user's password, even Administrator
adminGroup = "Domain Admins"
'resetGroup can only reset the passwords of the user group defined in usersGroup
resetGroup = "SSPR_PasswordChangers"
'-----------------
'Is semiSecretAnswer revealed on the user details page?
'True shows the answer, False requires the answer to be typed in for verification
showSemiSecretAnswer = False
'-----------------
'resetGroup can see users semi-secret question and answer. True or False?
resetGroupSemiSecret = False
'-----------------
'Active Directory Group whose passwords can be reset by resetGroup
usersGroup = "Students"
'-----------------
'Question array
'These are the security questions available to users
'Each question MUST have a unique sequential number, and the total number of questions must be listed in Dim QuestionArray()
Dim QuestionArray(11)
QuestionArray(0)="What is your Mother's maiden name?"
QuestionArray(1)="What is the last name of your favorite school teacher?"
QuestionArray(2)="What is the name of your favorite sports team?" 
QuestionArray(3)="What is the name of your favorite singer or band?" 
QuestionArray(4)="What is the name of your favorite television series?" 
QuestionArray(5)="What is the name of your favorite restaurant?" 
QuestionArray(6)="What is the name of your favorite movie?" 
QuestionArray(7)="What is the name of your favorite song?" 
QuestionArray(8)="What is the furthest place to which you have traveled?" 
QuestionArray(9)="What is the name of your favorite actor or actress?" 
QuestionArray(10)="Who is your personal hero?" 
QuestionArray(11)="What is your favorite hobby?"
'-----------------
End If
' End of user configuration
'Do not change the following lines
%>
<!--#include virtual="/iniread.asp"-->
<%
ImpersonateUser = ReadINI(CredFile, "Main", "ImpersonateUser")
ImpersonateUserPass = ReadINI(CredFile, "Main", "ImpersonateUserPass")
%>