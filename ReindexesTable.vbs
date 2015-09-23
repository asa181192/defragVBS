
'-------------------------------------------------------------------------------
'Set myLog = objfso.createtextfile("C:\Users\asantiago\Documents\VBSSCRIPTS\log.log",true)
'Const DB_CONNECT_STRING = "Provider=SQLOLEDB.1;Data Source=BACK-SQL1;Initial Catalog=TCS_TEST;Integrated Security=SSPI"
Set objfso = createobject("scripting.filesystemobject")
strPath = "C:\Users\asantiago\Documents\VBSSCRIPTS" 'Path para generar el log file en caso de que no exista

IF objfso.FileExists(strPath&"\log.log") = FALSE  THEN'Verifica si existe un archivo de log , en caso de que no el archivo se crea 
	
	call objfso.createtextfile(strPath & "\log.log",true) 
	 
END IF
 strPath = strPath&"\log.log"
 
Call WriteLogFileLine(strPath,"----------------------------Auto Generate log---------------------------")
On Error Resume NEXT
	Const DB_CONNECT_STRING = "Provider=SQLOLEDB.1;Data Source=BACK-SQL1;Initial Catalog=GRALTEST;Integrated Security=SSPI"	
	Set myConn = CreateObject("ADODB.Connection")
	Set AdRec = CreateObject("ADODB.Recordset")
	Set myCommand = CreateObject("ADODB.Command" )
	myConn.Open DB_CONNECT_STRING
	SQL = "select * from REINDEXESTABLE WHERE frag_in_percent >= 30 "
	AdRec.Open SQL, myConn
	
IF Err.Number <> 0 THEN 
	Call WriteLogFileLine(strPath,Err.Number &" "& Err.Description)
	myConn.Close 
	ELSE	
		Call WriteLogFileLine(strPath,"Conexion correcta ,conexion:"&DB_CONNECT_STRING)
		IF AdRec.EOF THEN 
			Call WriteLogFileLine(strPath,"No hay registros en la consulta")
			ELSE
			On Error Resume NEXT 
				Do While NOT AdRec.EOF
				IndexName = AdRec("INDEX_NAME")
				schemaname = AdRec("SCHEMA_NAME")
				objectname = AdRec("OBJECT_NAME")
				'wscript.echo IndexName
				Call reindex(IndexName,schemaname,objectname)
				AdRec.MoveNext
				LOOP
			IF Err.Number <> 0 THEN 
				Call WriteLogFileLine(strPath,Err.Number &" "& Err.Description)
				myConn.Close 
			END IF 

		END IF
END IF

myConn.Close 


'------------------------------------------------------------------------------- 
function reindex (indexname,schemaname,objectname)
	
	On Error Resume NEXT 
	'myConn.Open DB_CONNECT_STRING
	Set myCommand.ActiveConnection = myConn
	'myCommand.CommandText = "ALTER INDEX ["& indexname &"] ON ["& schemaname &"].["& objectname &"] REBUILD WITH (ONLINE=ON)"
	myCommand.CommandText = "INSERT INTO prueba (name) VALUES ('blabla')"
	myCommand.Execute
	IF Err.Number <> 0 THEN 
		Call WriteLogFileLine(strPath,Err.Number &" "& Err.Description)
		ELSE 
		Call WriteLogFileLine(strPath,"Execute query realizado correctamente, Query: "&myCommand.CommandText)
	END IF 
	'myConn.Close
End Function 



'----------------------------------------------------------------------------------

Function WriteLogFileLine(sLogFileName,sLogFileLine)
    dateStamp = Now() 

    Set objFsoLog = CreateObject("Scripting.FileSystemObject")
    Set logOutput = objfso.OpenTextFile(sLogFileName, 8, True)

    logOutput.WriteLine(cstr(dateStamp) + " -" + vbTab + sLogFileLine)
    logOutput.Close

    Set logOutput = Nothing
    Set objFsoLog = Nothing

End Function

