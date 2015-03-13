  _____      _  __    _____                 _
 / ____|    | |/ _|  / ____|               (_)
| (___   ___| | |_  | (___   ___ _ ____   ___  ___ ___
 \___ \ / _ \ |  _|  \___ \ / _ \ '__\ \ / / |/ __/ _ \
 ____) |  __/ | |    ____) |  __/ |   \ V /| | (_|  __/
|_____/ \___|_|_|   |_____/ \___|_|    \_/ |_|\___\___|
 _____                                    _   _____                _
|  __ \                                  | | |  __ \              | |
| |__) |_ _ ___ _____      _____  _ __ __| | | |__) |___  ___  ___| |_
|  ___/ _` / __/ __\ \ /\ / / _ \| '__/ _` | |  _  // _ \/ __|/ _ \ __|
| |  | (_| \__ \__ \\ V  V / (_) | | | (_| | | | \ \  __/\__ \  __/ |_
|_|   \__,_|___/___/ \_/\_/ \___/|_|  \__,_| |_|  \_\___||___/\___|\__|
      ___   ____   ___
     |__ \ |___ \ / _ \
__   __ ) |  __) | | | |
\ \ / // /  |__ <| | | |
 \ V // /_  ___) | |_| |
  \_/|____(_)___(_)___/

--------------------------------------------------------------------------------

    Copyright 2006 - 2015 Ben "Plexer" Norcutt and Alex "Irazmus" Kitching

    Self Service Password Reset is free software; you can redistribute it
    and/or modify it under the terms of the GNU General Public License as
    published by the Free Software Foundation; either version 2 of the
    License, or (at your option) any later version.

    Self Service Password Reset is distributed in the hope that it will
    be useful, but WITHOUT ANY WARRANTY; without even the implied warranty
    of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Self Service Password Reset; if not, write to the Free Software
    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

--------------------------------------------------------------------------------
** Requirements **
--------------------------------------------------------------------------------
* Active Directory
* IIS (Internet Information Services)
* ASP (Active Server Pages)
* ADODB

** Tested on **
--------------------------------------------------------------------------------
* Windows Server 2003 Standard Edition 
* Windows Server 2003 Standard Edition SP1
* Windows Server 2003 Standard Edition R2
* IIS 6.0

** ToDo **
--------------------------------------------------------------------------------
* Migrate from WinNT connector to LDAP connector
* Change ResetAttempts=0 to mean no limit rather than break SSPR as at present

** Change Log v2.3.0 **
--------------------------------------------------------------------------------
* Implement user search
* Add GLPI auto logging

--------------------------------------------------------------------------------
** Change Log v2.2.1 **
--------------------------------------------------------------------------------
* Added admin reset counter & last admin reset date to user details in admin console
* Added new table userdata to database to record last admin reset date and total number of admin resets per user
* Added Password Last Set to administration console

** Change Log v2.2 **
--------------------------------------------------------------------------------
* Modified admin\auth.asp to convert all group names to lower case to prevent case sensistive errors.
* Fixed problem where password cannot be changed if IIS is not running on a Domain Controller
* Added impersonation settings to config.asp for use with above fix
* Added LoginAdmin.dll for use with above fix
* Corrected error where update_details.asp was ignoring setting of HashAnswers
* Replaced all occurances of Response.Redirect "/register" with Response.Redirect "/register/register.asp"
* Removed evil hacks from reset_pass.asp and adminreset.asp, replaced with proper error checking

** Change Log v2.1 **
--------------------------------------------------------------------------------
* Added option for user answers to be held in plain text rather than hashed

** Change Log v2.0 **
--------------------------------------------------------------------------------
* Added update page to allow users to delete themselves from the database so they can register a new set of questions
* Added AD error handling for user lookup and password reset
* Added administration option to remove user from SSPR database allowing user to change their registered answers
* Rebuilt all files to use a templating system for easy reskinning and to reduce code repetition
* Created user admin section
* Expanded reset options. Administrator can now choose between 5 password options:
1) Static password used for all users
2) CSV file containing user/password pairs, on reset script looks up username and sets password to corresponding value
3) A random alphanumeric string of fixed or variable length
4) A password chosen at random from a TXT file of possibilities
5) A password chosen by the user
* Moved all implementation specific variables into config.asp for ease of setup
* Script now automatically expires users password upon reset, forcing user to change it on next logon (Unless "Password never expires" is set)
* Added check to reset.asp to output friendly error if user is not registered
* Added checks to prevent users registering more than once
* Added server-side form validation checking data is in valid format
* Added reset attempt counter with lockout after certain number of failed reset atempts
* Added client-side form validation for unselected/unanswered questions, duplicate questions, duplicate answers
* Replaced front end

** Quick Install **
--------------------------------------------------------------------------------
Assuming you have a default IIS install with no sites configured:
1) Extract the contents of this archive to C:\Inetpub

2) Register LoginAdmin.dll by running "regsvr32.exe c:\Inetpub\LoginAdmin.dll" from a command prompt on the Web server.

3) Set NTFS access rights for databases directory to Modify for all authenticated users

4) Open the IIS MMC, set the register, update and admin directory to Integrated Authentication and the reset directory to Anonymous Access

5) Still in the IIS MMC add index.asp as the default content page for update, admin and reset directory and register.asp as the default content page for the register directory

6) Still in the IIS MMC set Active Server Pages to Allow under Web Service Extensions

7) Copy C:\Inetpub\wwwroot\config.asp.dist to C:\Inetpub\wwwroot\config.asp

8) Open C:\Inetpub\wwwroot\config.asp with notepad and change settings as required. The only required change is to set your FQDN

9) Copy C:\inetpub\cred.ini.dist to C:\Inetpub\cred.ini

10) Open C:\Inetpub\cred.ini with notepad and enter the logon credentials of a user with write access to Active Directory (Suggest a new account called SSPR_MACHINENAME with a long password which doesn't expire). Do not remove the end section header

11) Set users IE Homepage to http://yourserver/register/register.asp

12) Create a highly restricted passwordless user account called resetpassword whose user interface is set to "C:\Program Files\Internet Explorer\iexplore.exe -k http://yourserver/reset/index.asp"

13) Provide a link for you users to http://yourserver/update/index.asp so they can update their details

That's it

--------------------------------------------------------------------------------

If you have any problems or would like to leave feedback, go to http://www.edugeek.net/forums/showthread.php?t=2022