<div align="center">

## SQL and ASP


</div>

### Description

SQL and ASP:

Here I have listed same important SQL-Statement.

and same functions for connect to Database.

By my samples I don't use:

follow ADO Command:

-add

-update

...

only SQL (is faster and better to read)

I dont't check why so much programmer use not the SQL-Staetment
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hohl David](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hohl-david.md)
**Level**          |Beginner
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__4-35.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hohl-david-sql-and-asp__4-7973/archive/master.zip)





### Source Code

```
'SELECT with Array
Dim arrstrZipcode()
Dim i
Set rec = Server.CreateObject("ADODB.Recordset")
strSQL = "SELECT zipcode FROM ZIPCODES WHERE Countrycode = 'AUT'"
set rec = objConn.Execute(strSQL)
if not rec.eof then
	while not rec.eof
		redim preserve arrstrZipcode(i)
		arrstrZipcode(i) = rec.fields("zipcode")
		rec.movenext
		i = i + 1
	wend
else
	strError = "Nothing found"
end if
' ——————————————————————————————————————
'UPDATE
strFirstname = ReplaceUnletters(Request.form("txtFirstname"))
iAdressID = Request.Form("txtAdressID")
Set rec = Server.CreateObject("ADODB.Recordset")
strSQL = "Update ADDRESS SET " & _
	 "Firstname = '" & strFirstname & "' " & _
	 "WHERE ZipcodeID = " & iAdressID
set rec = objConn.Execute(strSQL)
' ——————————————————————————————————————
'INSERT
strFirstname = ReplaceUnletters(Request.form("txtFirstname"))
strLastname = ReplaceUnletters(Request.form("txtFirstname"))
Set rec = Server.CreateObject("ADODB.Recordset")
strSQL = "INSERT INTO ADDRESS (Firstname, Lastname) VALUES ( " & _
	 "'" & strFirstname & "', " & _
	 "'" & strLastname & "')"
set rec = objConn.Execute(strSQL)
' ——————————————————————————————————————
'DELETE
iAddressID = Request.querystring("AdID")
Set rec = Server.CreateObject("ADODB.Recordset")
strSQL = "DELETE FROM ADDRESS WHERE AddressID = " & iAddressID
set rec = objConn.Execute(strSQL)
' ——————————————————————————————————————
' FUNCTIONS
' ——————————————————————————————————————
function OpenDatabase()
Dim strDatabaseIP, strPW, strUID
	strDatabaseIP = "150.150.100.210"
	strPW = ""
	strUID = "sa"
    'SQL-Server
	strConn = "Provider=SQLOLEDB.1;SERVER=" & strDatabaseIP & ";DATABASE=ADDRESS;UID=" & strUID & ";PWD=" & strPW
'MY SQL
strConn = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strDatabaseIP & ";DATABASE=Address;UID=" & strUID & ";PWD=" & strPW & ";OPTION=35;"
	set objConn = server.CreateObject("ADODB.Connection")
	objConn.Open(strConn)
end function
function CloseDatabase()
	objConn.close
	Set objConn = Nothing
end function
' ——————————————————————————————————————
Public Function ReplaceUnLetters(ByVal strReplace)
	strReplace = Replace(strReplace, "'", "''")
	ReplaceUnLetters = strReplace
End Function
' ——————————————————————————————————————
Function ReplaceNonSpace(strText)
'Here I want replace HTMLSPACE WITH normal space, but the insert into the Planet Source code replace the HTML TAG excuse
	ReplaceNonSpace = Replace(strText, " ", " " )
End Function
```

