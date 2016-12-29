call main

function Connection_and_RecordSet()
	Set Connection = CreateObject("ADODB.Connection")
	Set Recordset = CreateObject("ADODB.Recordset")
	
	' configure ODBC and input your config here'
	Connection.Open "DSN=SQLServer;UID=sa;PWD=sasasa#;Database=example"
	' git update-index --assume-unchanged main.vbs after you change you pass and UID
	' that is needed for not uploading changed config to git, secuirity mother Russia

	' type_desc = folder, name = file name, and text = file content
	SQL = "select type_desc, o.name, replace(replace(replace(C.B,'&#x0D;',''),'&lt;', '<'),'&gt;', '>') text from sys.all_objects o "&_
	  "outer apply (select (select s.text+char(13) from sys.syscomments s where s.id = o.object_id for xml path('')) B ) C "&_
	  "where type_desc is not null and C.B <> '(server internal)' UNION ALL "&_
	  "select 'TABLES',t.name, "&_
	  "replace((select COLUMN_NAME+' '+DATA_TYPE+isNULL(' ('+CAST(CHARACTER_MAXIMUM_LENGTH as VARCHAR)+')','')+ISNULL(' DEFAULT '+COLUMN_DEFAULT+';',';')+CHAR(13) "&_
	  "from INFORMATION_SCHEMA.COLUMNS d where d.TABLE_NAME = t.name for xml path('')),'&#x0D;',CHAR(13)) text from sys.tables t"

	Recordset.Open SQL, Connection
	Set Connection_and_RecordSet = Recordset
end Function 
 

sub main()
	Set fso = CreateObject("Scripting.FileSystemObject")
	current_dir  = fso.GetAbsolutePathName(".")

	set RS = Connection_and_RecordSet()
	Do While NOT RS.Eof
		write_to_file current_dir, RS("type_desc"), RS("name"), RS("text")
		RS.MoveNext     
	Loop
	commit_and_push_to_git
end sub

sub write_to_file(current_path, folder, name, text)
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	strDirectory = current_path & "\" & folder
	If objFSO.FolderExists(strDirectory) Then
   		Set objFolder = objFSO.GetFolder(strDirectory)
	Else
   		Set objFolder = objFSO.CreateFolder(strDirectory)
	End If
	outFile="\"+name+".sql"
	Set objFile = CreateObject("ADODB.Stream")
	objFile.Type = 2 'Specify stream type - we want To save text/string data.
	objFile.Charset = "utf-8" 'Specify charset For the source text data.
	objFile.Open 
	objFile.WriteText text
	objFile.SaveToFile strDirectory&outFile, 2
end sub

sub commit_and_push_to_git()
	Set oShell = WScript.CreateObject("WSCript.shell")
	oShell.run "cmd.exe /C git add . "
	WScript.sleep 5000
	oShell.run "cmd.exe /C git commit -m """+Replace(Now()," "," ")+""""
	WScript.sleep 5000
	oShell.run "cmd.exe /C git push -f"
end sub
