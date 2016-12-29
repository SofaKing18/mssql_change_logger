call main

function Connection_and_RecordSet()
	Set Connection = CreateObject("ADODB.Connection")
	Set Recordset = CreateObject("ADODB.Recordset")
	
	' configure ODBC and input your config here'
	Connection.Open "DSN=SQLServer;UID=sa;PWD=sasasa;Database=example"
	' git update-index --assume-unchanged main.vbs after you change you pass and UID
	' that is needed for not uploading changed config to git, secuirity mother Russia

	' type_desc = folder, name = file name, and text = file content
	SQL = PROCEDURES_AND_ETC_SQL()&" UNION ALL "&TABLES_SQL()&" UNION ALL "&INDEXES_SQL()

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


function TABLES_SQL()
TABLES_SQL=_
	"select 'TABLES',t.name, "&_
	"replace((select COLUMN_NAME+' '+DATA_TYPE+isNULL(' ('+CAST(CHARACTER_MAXIMUM_LENGTH as VARCHAR)+')','')+ISNULL(' DEFAULT '+COLUMN_DEFAULT+';',';')+CHAR(13) "&_
	"from INFORMATION_SCHEMA.COLUMNS d where d.TABLE_NAME = t.name for xml path('')),'&#x0D;',CHAR(13)) text from sys.tables t"
end Function 

function PROCEDURES_AND_ETC_SQL()
PROCEDURES_AND_ETC_SQL =_
	"select type_desc, o.name, replace(replace(replace(C.B,'&#x0D;',''),'&lt;', '<'),'&gt;', '>') text from sys.all_objects o "&_
	"outer apply (select (select s.text+char(13) from sys.syscomments s where s.id = o.object_id for xml path('')) B ) C "&_
	"where type_desc is not null and C.B <> '(server internal)'"
End Function

function INDEXES_SQL()
' https://gallery.technet.microsoft.com/scriptcenter/sql-server-generate-index-fa790441
' thx for this script :)
INDEXES_SQL =_ 
"SELECT 'INDEXES','['+T.name+'] '+I.name,' CREATE ' + "+_
"  CASE WHEN I.is_unique = 1 THEN ' UNIQUE ' ELSE '' END  +  "+_
"  I.type_desc COLLATE DATABASE_DEFAULT +' INDEX ' +   "+_
"  I.name  + ' ON '  +  "+_
"  Schema_name(T.Schema_id)+'.'+T.name + ' ( ' + "+_
"  KeyColumns + ' )  ' + "+_
"  ISNULL(' INCLUDE ('+IncludedColumns+' ) ','') + "+_
"  ISNULL(' WHERE  '+I.Filter_definition,'') + ' WITH ( ' + "+_
"  CASE WHEN I.is_padded = 1 THEN ' PAD_INDEX = ON ' ELSE ' PAD_INDEX = OFF ' END + ','  + "+_
"  'FILLFACTOR = '+CONVERT(CHAR(5),CASE WHEN I.Fill_factor = 0 THEN 100 ELSE I.Fill_factor END) + ','  + "+_
"  'SORT_IN_TEMPDB = OFF '  + ','  + "+_
"  CASE WHEN I.ignore_dup_key = 1 THEN ' IGNORE_DUP_KEY = ON ' ELSE ' IGNORE_DUP_KEY = OFF ' END + ','  + "+_
"  CASE WHEN ST.no_recompute = 0 THEN ' STATISTICS_NORECOMPUTE = OFF ' ELSE ' STATISTICS_NORECOMPUTE = ON ' END + ','  + "+_
"  ' DROP_EXISTING = ON '  + ','  + "+_
"  ' ONLINE = OFF '  + ','  + "+_
"  CASE WHEN I.allow_row_locks = 1 THEN ' ALLOW_ROW_LOCKS = ON ' ELSE ' ALLOW_ROW_LOCKS = OFF ' END + ','  + "+_
"  CASE WHEN I.allow_page_locks = 1 THEN ' ALLOW_PAGE_LOCKS = ON ' ELSE ' ALLOW_PAGE_LOCKS = OFF ' END  + ' ) ON [' + "+_
"  DS.name + ' ] '  [CreateIndexScript] "+_
"FROM sys.indexes I   "+_
"JOIN sys.tables T ON T.Object_id = I.Object_id    "+_
"JOIN sys.sysindexes SI ON I.Object_id = SI.id AND I.index_id = SI.indid   "+_
"JOIN (SELECT * FROM (  "+_
"  SELECT IC2.object_id , IC2.index_id ,  "+_
"    STUFF((SELECT ' , ' + C.name + CASE WHEN MAX(CONVERT(INT,IC1.is_descending_key)) = 1 THEN ' DESC ' ELSE ' ASC ' END "+_
"    FROM sys.index_columns IC1  "+_
"    JOIN Sys.columns C   "+_
"    ON C.object_id = IC1.object_id   "+_
"    AND C.column_id = IC1.column_id   "+_
"    AND IC1.is_included_column = 0  "+_
"    WHERE IC1.object_id = IC2.object_id   "+_
"    AND IC1.index_id = IC2.index_id   "+_
"    GROUP BY IC1.object_id,C.name,index_id  "+_
"    ORDER BY MAX(IC1.key_ordinal)  "+_
"    FOR XML PATH('')), 1, 2, '') KeyColumns   "+_
"    FROM sys.index_columns IC2   "+_
"    GROUP BY IC2.object_id ,IC2.index_id) tmp3 )tmp4   "+_
"  ON I.object_id = tmp4.object_id AND I.Index_id = tmp4.index_id  "+_
"JOIN sys.stats ST ON ST.object_id = I.object_id AND ST.stats_id = I.index_id   "+_
"JOIN sys.data_spaces DS ON I.data_space_id=DS.data_space_id   "+_
"JOIN sys.filegroups FG ON I.data_space_id=FG.data_space_id   "+_
"LEFT JOIN (SELECT * FROM (   "+_
"  SELECT IC2.object_id , IC2.index_id ,   "+_
"    STUFF((SELECT ' , ' + C.name  "+_
"    FROM sys.index_columns IC1   "+_
"    JOIN Sys.columns C    "+_
"       ON C.object_id = IC1.object_id    "+_
"       AND C.column_id = IC1.column_id    "+_
"       AND IC1.is_included_column = 1   "+_
"    WHERE IC1.object_id = IC2.object_id    "+_
"      AND IC1.index_id = IC2.index_id    "+_
"    GROUP BY IC1.object_id,C.name,index_id   "+_
"      FOR XML PATH('')), 1, 2, '') IncludedColumns    "+_
"  FROM sys.index_columns IC2    "+_
"  GROUP BY IC2.object_id ,IC2.index_id) tmp1   "+_
"  WHERE IncludedColumns IS NOT NULL ) tmp2    "+_
"ON tmp2.object_id = I.object_id AND tmp2.index_id = I.index_id   "+_
"WHERE I.is_primary_key = 0 AND I.is_unique_constraint = 0 "
end Function