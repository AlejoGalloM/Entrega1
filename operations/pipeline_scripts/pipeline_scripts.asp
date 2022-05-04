<OBJECT RUNAT=server PROGID=ADODB.Connection id=objConnpipelineDB VIEWASTEXT> </OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Connection id=objConnpipelineDB_AUDIO VIEWASTEXT> </OBJECT>

<SCRIPT LANGUAGE=vbscript RUNAT=Server>
'===================================================================================
'File Name:		pipeline_scripts.asp
'Path:				operations/_pipeline_scripts/
'Created By:
'Last Modified:	A. Orozco 2001/09/24
'						Guillermo Aristizabal 2001/10/11
'Parameters:		None
'Returns:			
'Additional Information:	Several functions used by manacha
'									Last modification: added write_sp_log function
'===================================================================================


'==========================================================================
'<I&T - DMPC 2009/12/05 - Proyecto Reforma Financiera>
' Realiza la invocaci?n del web service SkCo.Transactions.FSL para obtener la informaci?n de Transactions_History
'==========================================================================
function InvokeTransactionsWS(url)
dim xmlDOC
Dim bOK
Dim HTTP
dim accion

Set HTTP = CreateObject("MSXML2.XMLHTTP")
Set xmlDOC =CreateObject("MSXML.DOMDocument")
xmlDOC.Async=False

accion = Application.Contents("WSTransactionHistory") & "/" & url 
HTTP.Open "GET", accion , false
HTTP.Send(null)
bOK = xmlDOC.load(HTTP.responseXML)

if Not bOK then
response.Write HTTP.responseXML
exit Function
end if
    
Dim result
result=""

result = xmlDOC.Text

InvokeTransactionsWS = result
end function
'=======================================================================================

'=======================================================================================
'Imported scipts from pipeline
'=======================================================================================
'<I&T - WTG> Admisi?n de caracteres alfanumericos Pasaporte
Function ReadNumber(strNumber)
	Dim number
	Dim i
	For i = 1 To Len(strNumber)
		If Mid(strNumber,i,1) >= "0" And Mid(strNumber,i,1) <= "9" Then
			number = number & Mid(strNumber,i,1)
		End If
	Next
	ReadNumber = number
End Function
'</I&T>

Function sp_trim_all(strToTrim)
	sp_trim_all = LTrim(RTrim(CStr(strToTrim)))
End Function

Function autorizarMenu(pagina, caracter)

	If pagina = -1 And Caracter = -1 Then
		autorizarMenu = 1 
		Exit Function
	End If  

	mipermiso = session("sp_permisos")
	If Len(Session.Contents("sp_permisos")) = 0 Then
		Response.Redirect "../error_pages/timeout.asp" 
	End If

	'Check Login status
	If Len(mipermiso) = 0 Then
		Response.Redirect "default.asp"
	End If

	'caracter exista
	If Len(mipermiso) < caracter Then
		Response.Redirect "p_NoPermiso.asp"
	End If
	
	'pagina=pagina-1
	mipermiso = Mid(mipermiso,caracter,1)
	'Check if the user has authorization
	If (Asc(mipermiso) and 2^pagina)<>2^pagina Then
		autorizarMenu = 0 'Option Unauthorized User
	Else
		autorizarMenu = 1 'Option Authorized User
	End If  
End Function

Sub Authorize(pagina, caracter)
	Dim mipermiso

	mipermiso = Session("sp_permisos")
	'Time Excess
	If Len(Session.Contents("sp_permisos")) = 0 Then
		Response.Redirect Application("TimeoutURL")
	End If
	'Check login status
	If Len(mipermiso) = 0 Then
		Response.Redirect "default.asp"
	End If
	'si existe el caracter
	If Len(mipermiso) < caracter Then
		Response.Redirect "p_NoPermiso.asp"
	End If  
	'pagina=pagina-1
	mipermiso = Mid(mipermiso, caracter, 1)
	'verificar si tiene permiso a la pagina
	If (Asc(mipermiso) And 2^pagina) <> 2^pagina Then
		Response.Redirect Application("UnauthorizedURL") & "?permiso=" & pagina & "/" & caracter & "/mperm=" & Asc(mipermiso)
	End If
End Sub

'p1 es fecha en afiliacion 
'p2 es hora
'p3 es consecutivo
'p4 es unidad cuando haya, 0 cuando no

Sub writelog2(pagina, sp, producto,error,plan,texto,p1,p2,p3,p4)
	Dim sSql, rlog
	sSql = "exec sppl_putlog " & CStr(Session.sessionID) & "," & CStr(pagina) & ",'" & _
	Request.ServerVariables("REMOTE_ADDR") & "','" & sp & "'," & CStr(session("sp_idworker")) & "," & CStr(0) & ",'" & _
	producto & "'," & CStr(0) & "," & CStr(error) & ",'" & plan & "'" & ",'" & texto & "'," & CStr(p1) & "," & CStr(p2) & "," & _
	CStr(p3) & "," & CStr(p4)
	Set objConn = GetConnPipelineDb
	objConn.Execute sSql
	CloseConnPipelineDb
End Sub

Sub write_sp_log(connection, page_id, sp, contract, product, plan, client_id, error, conf_num, text)
	Dim strSQL
	text = "SP Log - " & text
	If IsNull(Session("sp_idworker")) Or Session("sp_idworker") = "" Then
		Session("sp_idworker") = "0"
	End If
	
		
	strSQL = "spsp_PutLog " & _
	Session.SessionID & ", " & _
	page_id & ", " & _
	"'" & Request.ServerVariables("REMOTE_ADDR") & "', " & _
	"'" & sp & "', " & _
	Session("sp_idworker") & ", " & _
	contract & ", " & _
	"'" & product & "', " & _
	client_id & ", " & _
	error & ", " & _
	"'" & plan & "', " & _
	"'" & conf_num & "', " & _
	"'" & text & "'"
	connection.Execute strSQL
End Sub

'New method
Sub write_sp_log_Figue(connection, page_id, sp, contract, product, plan, client_id, error, conf_num, text)
	Dim strSQL
	text = "SP Log - " & text
	If IsNull(Session("sp_idworker")) Or Session("sp_idworker") = "" Then
		Session("sp_idworker") = "0"
	End If
	
    dim command
    dim paramSession
    dim paramPage
    dim paramWorker
    dim paramContract
    dim paramClient
    dim paramError

    Set command = server.CreateObject("ADODB.Command")
    command.CommandText = "spsp_PutLog"
    command.CommandType = 4 'adCmdStoredProc
    command.ActiveConnection = connection

    set paramSession = command.CreateParameter("@idSession", 131, 1, 9, Session.SessionID) '131=adNumeric, 1=adInput
    paramSession.NumericScale = 0
    paramSession.Precision = 18
    command.Parameters.Append(paramSession)

    set paramPage = command.CreateParameter("@Page", 131, 1, 9, page_id)
    paramPage.NumericScale = 0
    paramPage.Precision = 18
    command.Parameters.Append(paramPage)

    command.Parameters.Append(command.CreateParameter("@IpClient", 200, 1, 15, Request.ServerVariables("REMOTE_ADDR"))) '200=adVarChar, 1=adInput
    command.Parameters.Append(command.CreateParameter("@XSProcedure", 200, 1, 50, sp)) '200=adVarChar, 1=adInput

    set paramWorker = command.CreateParameter("@idWorker", 131, 1, 0, Session("sp_idworker"))
    paramWorker.NumericScale = 0
    paramWorker.Precision = 18
    command.Parameters.Append(paramWorker)

    set paramContract = command.CreateParameter("@Contrato", 131, 1, 0, contract)
    paramContract.NumericScale = 0
    paramContract.Precision = 18
    command.Parameters.Append(paramContract)

    command.Parameters.Append(command.CreateParameter("@producto", 200, 1, 8, product)) '200=adVarChar, 1=adInput

    set paramClient = command.CreateParameter("@NroDocCli", 131, 1, 0, client_id)
    paramClient.NumericScale = 0
    paramClient.Precision = 18
    command.Parameters.Append(paramClient)

    set paramError = command.CreateParameter("@idError", 131, 1, 0, error)
    paramError.NumericScale = 0
    paramError.Precision = 18
    command.Parameters.Append(paramError)

    command.Parameters.Append(command.CreateParameter("@Plan", 200, 1, 8, plan)) '200=adVarChar, 1=adInput
    command.Parameters.Append(command.CreateParameter("@ConfirmationNumber", 200, 1, 17, conf_num)) '200=adVarChar, 1=adInput
    command.Parameters.Append(command.CreateParameter("@TextoBuscar", 200, 1, 255, text)) '200=adVarChar, 1=adInput

    command.Execute

End Sub
'Fin Alejandro Figueroa

'M?todo para enviar logs a Servicio REST
Sub write_dataLog(typeMessage,componentId,processInfo, userAffected, spaceAffected, valueData, valueNewData, internalProcess,exitDate)
	Dim jsonText, authorization
	Dim processDate, processDateFormat, dayDate, timeDate 
	Dim WMI  
	Dim Nads 
	Dim nad
	Dim valueNew, value, message, value64, urlService,urlServiceAuth, token
	Dim HTTP1, HTTPA
	Set HTTP1 = CreateObject("MSXML2.XMLHTTP")
	Set HTTPA = CreateObject("MSXML2.XMLHTTP")

	Set WMI = GetObject("winmgmts:\\.\root\cimv2")
	Set Nads = WMI.ExecQuery("Select * from Win32_NetworkAdapter where physicaladapter=true")

	dayDate= Date
	timeDate= Time
	processDate= year(dayDate)&"-"&month(dayDate)&"-"&Day(dayDate)&" "&timeDate

	Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"

	IF valueNewData="null" then
		valueNew="null"
	else
		valueNew = "{" & """accessDate"": " & """" & Session("accessDate") & """," & """userAffected"": " & """" & userAffected & """," & """process"": " & """" & processInfo & """," & """dateProcess"": " & """" & processDate & """," & """internalProcess"": " & """" & internalProcess & """," & """spaceAffected"": " & """" & spaceAffected & """," & 	"""value"": " & """" & value & """," & """valueNew"": " & """" & valueNewData & """," & """user"": " & """" & Request.Cookies("sp_idworker") & """," & """ipPublic"": " & """" & Session("ipPublica") & """," & """browser"": " & """" & Request.Cookies("browser") & """," & """exitDate"": " & """" & exitDate & """}"
		oNode.nodeTypedValue =Stream_StringToBinary(valueNew)
		valueNew=oNode.text		
	end IF

	IF valueData="N/A" then
		value="N/A"
	ELSE
		value="{"&valueData&"}"		
	end IF

	value64= "{" & """accessDate"": " & """" & Session("accessDate") & """," & """userAffected"": " & """" & userAffected & """," & """process"": " & """" & processInfo & """," & """dateProcess"": " & """" & processDate & """," & """internalProcess"": " & """" & internalProcess & """," & """spaceAffected"": " & """" & spaceAffected & """," & """value"": " & """" & value & """," & """valueNew"": " & """" & valueNewData & """," & """user"": " & """" & Request.Cookies("sp_idworker") & """," & """ipPublic"": " & """" & Session("ipPublica") & """," &  """browser"": " & """" & Request.Cookies("browser") & """," & """exitDate"": " & """" & exitDate & """}"
	oNode.nodeTypedValue =Stream_StringToBinary(value64)
	value64=oNode.text

	if spaceAffected="" then
		spaceAffected="null"
	end if

	if userAffected=Session("sp_miLogin") or userAffected="" then
		userAffected="null"
	end if

	if typeMessage="200 OK" then
		message="Info"
	else
		if typeMessage="Warning" then
			message="Warning"
		else
			message="Error"
		end if
	end if
	
	jsonText = "{" & """MessageType"": " & """" & message & """," & """ComponentId"": " & """Pipeline 2 -" & componentId & """," & """InfoDate"": " & """" & Session("accessDate") & """," & """InfoProcess"": " & """" & processInfo & """," & """InfoValue"": "& """" &value64 & """," &"""InfoValueNew"": " & """"& valueNew & """," & """InfoUser"": " & """" & Request.Cookies("sp_idworker") & """," & """InfoPublicIp"": " & """" & Session("ipPublica") & """," & """InfoBrowser"": " & """" & Request.Cookies("browser") & """" & "}"
	
	authorization = Application.Contents("LOGSAUTBODY")
	
	urlService = Application.Contents("LOGSURL")
	urlServiceAuth = Application.Contents("LOGSURLAUTH")

	on error resume next
	HTTPA.Open "POST", urlServiceAuth, False
	If Err Then            'handle errors
	  Response.Write Err.Description & " [0x" & Hex(Err.Number) & "]"
	  WScript.Quit 1
	End If
	
	HTTPA.setRequestHeader "Content-Type", "application/json"
	HTTPA.send authorization
	
	if valueData<>"null" then
		token = Mid(HTTPA.responseText, 30, 280)
	end if
	
	

	On Error Goto 0        'disable error handling again
	
	 HTTP1.Open "POST", urlService, False
	 If Err Then            'handle errors
	   Response.Write Err.Description & " [0x" & Hex(Err.Number) & "]"
	   WScript.Quit 1
	 End If
	 HTTP1.setRequestHeader "Content-Type", "application/json"
	 HTTP1.setRequestHeader "Authorization", "Bearer "&token

	 HTTP1.send jsonText

	' On Error Goto 0        'disable error handling again
End Sub

Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Sub writelog(pagina, sp, producto,error,plan)
'	dim sSql
'	sSql = "exec sppl_putlog " & CStr(Session.sessionID) & "," & CStr(pagina) & ",'" & _
'	Request.ServerVariables("REMOTE_ADDR") & "','" & sp & "'," & CStr(session("sp_idworker")) & "," & _
'	CStr(Session.Contents("sp_ctr")) & ",'" & producto & "'," & CStr(session("sp_nrodocum")) & "," & CStr(error) & ",'" & plan & "'"
'	rlog.close()
'	rlog.setSQLText sSql
'	rlog.open()
'End Sub

Sub PM_transferTablesDesde()
	Dim i
	Dim a
	i = 0
	PM_rdFundContrato.moveFirst
	While Not PM_rdFundContrato.EOF
		i = i +1
		PM_renglonDesde  PM_rdFundContrato.fields.getValue("idFund") , i 
		PM_rdFundContrato.moveNext
	Wend
End Sub


Function PM_gNoAsset()
	Dim sSqlfc
	sSqlfc = "exec sppl_getfondos_contratos " & CStr(Session.Contents("sp_ctr"))
	PM_rdFundContrato.close()
	PM_rdFundContrato.setSQLText sSqlfc
	PM_rdFundContrato.open()
	PM_gNoAsset = True
	If PM_rdFundContrato.getCount() = 0 Then
		' error interno no puede suceder...
		'reinicio("contrato")
		PM_gNoAsset = False	
	End If
End Function 

Function PM_POrcenta (money)
	PM_Porcenta = FormatPercent(money, 2)
End Function

Function PM_Dinero (money)
	PM_Dinero = FormatCurrency(money, 2)
End Function

Function PM_EsPar (X)
	If ((Int(X / 2) ) * 2 ) = X Then
		PM_EsPar = 1
	Else
		PM_EsPar = 0
	End If
End Function

Function autorizarMn(pagina, caracter)
	Dim mipermiso
	
	autorizarMn = True
	If pagina = -1 And caracter = -1 Then
		Exit Function
	End If
	mipermiso = session("sp_permisos")
	If Len(mipermiso) < caracter Then
		autorizarMn = False
		Exit Function
	End If

	mipermiso = Mid(mipermiso, caracter, 1)
	If (Asc(mipermiso) And 2^pagina) <> 2^pagina Then
		autorizarMn = False
	End If  
End Function
'=======================================================================================
'End Of imported scipts from pipeline
'=======================================================================================

Function PasswordGenerator(ByVal lngLength)

	' Description: Generate a random password of 'user input' length
	' Parameters : lngLength - the length of the password to be generated
	' Returns    : String    - Randomly generated password
	' Created    : 2001/08/16 A. Orozco
	  
	Dim iChr
	Dim c
	Dim strResult
	Dim iAsc
	 
	Randomize Timer

	For c = 1 To lngLength

		' Randomly decide what set of ASCII chars we will use
		iAsc = Int(3 * Rnd + 1)

		'Randomly pick a char from the random set
		Select Case iAsc
			Case 1
				iChr = Int((Asc("Z") - Asc("A") + 1) * Rnd + Asc("A"))
			Case 2
				iChr = Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a"))
			Case 3
				iChr = Int((Asc("9") - Asc("0") + 1) * Rnd + Asc("0"))
			Case Else
				Err.Raise 20000, , "PasswordGenerator has a problem."
		End Select

		strResult = strResult & Chr(iChr)

	Next

	PasswordGenerator = strResult

End Function

Sub authorized(pagina, caracter)
dim mipermiso

  mipermiso=session("sp_profile")
  'exceso de tiempo
  if (len(Session.Contents("sp_profile")) = 0  ) then
	writelog3 pagina, 0, "","",1,1,"","Esmeralda: authorized : esmeraldascripts.asp : Exceso de tiempo"
    Response.Redirect "../error/critical_error_message.asp?mensaje=Su%20sesi?n%20ha%20sido%20terminada%20por%20exceso%20de%20tiempo%20inactivo.%20Por%20favor%20ingrese%20nuevamente.et"
  end if
  'verificar que este log  
  if len(mipermiso)=0 then
	writelog3 pagina, 0, "","",1,2,"","Esmeralda: authorized : esmeraldascripts.asp : Exceso de tiempo, mipermiso = 0"
    Response.Redirect "../error/critical_error_message.asp?mensaje=Su%20sesi?n%20ha%20sido%20terminada%20por%20exceso%20de%20tiempo%20inactivo.%20Por%20favor%20ingrese%20nuevamente.perm"
  end if
  'si existe el caracter
  if len(mipermiso)<caracter then
	writelog3 pagina, 0, "","",1,3,"","Esmeralda: authorized : esmeraldascripts.asp : error de aplicacion, bit no registrado para el usuario, debe observar si el error es en la pagina o en la informacion del usuario "
    Response.Redirect "../error/critical_error_message.asp?mensaje=Ud. no tiene autorizacion a esta opcion, su accion ha sido registrada en el log de auditoria."
  end if  
  'pagina=pagina-1
  mipermiso=Mid(mipermiso,caracter,1)
  'verificar si tiene permiso a la pagina
  if (Asc(mipermiso) and 2^pagina)<>2^pagina then
    writelog3 pagina, 0, "","",1,4,"","Esmeralda: authorized : esmeraldascripts.asp : intento de ingreso sin autorizacion."
    Response.Redirect "p_NoPermiso.asp?permiso="+ cstr(2^pagina)
  end if  
  'Verificar si puede ver el contrato
	If Request.QueryString("Contract") <> "" Then
	Dim aut
		For J = 0 To UBound(Session("sp_Contracts"),2)
			If Cstr(Request.QueryString("Contract")) = Cstr(Session("sp_Contracts")(0,J)) And Cstr(Trim(Request.QueryString("Product"))) = Cstr(Trim(Session("sp_Products")(0,J))) And Cstr(Request.QueryString("Plan")) = Cstr(Trim(Session("sp_Plans")(0,J))) Then
				aut = True
				Exit For
			Else
				aut =False
			End If
		Next
		If Not(aut) Then
			Response.Redirect "../error/critical_error_message.asp?mensaje=Ud.%20no%20tiene%20autorizacion%20a%20esta%20opcion,%20su%20accion%20ha%20sido%20registrada%20en%20el%20log%20de%20auditoria."
		End If
	End If
end sub

function getWholeSalerSocsArray(idWholeSaler)
			Dim strSql 'Stored procedures and SQL queries
			Dim objConn 'ADODB Connection
			Dim objRst 'ADODB Recordset			
			set objRst = Server.CreateObject("ADODB.Recordset")
			
			Set objConn = GetConnpipelineDB
			strSql = "spsp_GetWSSocieties '" & idWholeSaler & "'"
			objRst.Open strSql, objConn 
			If objRst.BOF And objRst.EOF Then
				getWholeSalerSocsArray = 0
			Else
				getWholeSalerSocsArray = objRst.GetRows()
			End If
			objRst.Close				
			set objConn = nothing
			set objRst = nothing
end function

function isContractInSocs(ContractSoc, arrSocs)
	Dim K
	if not IsArray(arrSocs) then 
		isContractInSocs = false
		exit function 
	end if		
	For K = 0 To UBound(arrSocs,2)
			If CStr(arrSocs(0,K)) = CStr(ContractSoc) Then
				isContractInSocs = True
				Exit For
			Else
				isContractInSocs = False
			End If
	Next
end function

Sub AuthorizeContractAccess(AccessLevel, idAgteLoggedIn, idAgteContract, idSocContract)
	dim isAuthorized 
	isAuthorized  = false
	'isAuthorized = true	
	'EXIT SUB

	Select Case AccessLevel
		Case 0 'Skandia
			isAuthorized  = true		
		Case 1 'WHOLE SALER		
			dim arrSocs
			arrSocs = getWholeSalerSocsArray(Session("sp_Idworker"))			
			isAuthorized = isContractInSocs(idSocContract,arrSocs)
		Case 2 'Partner FP, Fran. Worker
			if cstr(idSocContract) = cstr(Session("sp_idSoc")) then
				isAuthorized = true
			else
				isAuthorized = false
			end if
		Case 3 'FP
			if cstr(idAgteLoggedIn) = cstr(idAgteContract) then
				isAuthorized = true
			else
				isAuthorized = false
			end if	
		End Select
	if not isAuthorized then
		Response.Write "<br>idAgteContract = " 
		Response.Write idAgteContract
		Response.Write "<br>idSocContract = "
		Response.Write idSocContract
		Response.Write "<br>cstr(idAgteLoggedIn) : "
		Response.Write cstr(idAgteLoggedIn)
		Response.Write "<br>cstr(idAgteContract) : "
		Response.Write cstr(idAgteContract)
		Response.Write "<br><br><br>"
		Session.Contents("sp_permisos") = ""
		Session.Abandon		 
		Response.Write "<br>cstr(Session('sp_idSoc')) = "
		Response.Write cstr(Session("sp_idSoc"))
		Response.Write "SessionId"
		Response.Write Session.SessionID 
		Response.End
		'Response.Redirect Application("UnauthorizedURL")		  
	end if
end sub

Function GetConnpipelineDB
	If objConnpipelineDB.State = 0 Then
		objConnpipelineDB.ConnectionTimeout = Application("pipelineDB_ConnectionTimeout")
		objConnpipelineDB.CommandTimeout = Application("pipelineDB_CommandTimeout")
		objConnpipelineDB.CursorLocation = Application("pipelineDB_CursorLocation")
		objConnpipelineDB.Open Application("pipelineDB_ConnectionString"), Application("pipelineDB_RuntimeUserName"), Application("pipelineDB_RuntimePassword")
	End If
	Set GetConnpipelineDB = objConnpipelineDB
End Function

sub CloseConnpipelineDB
	If objConnpipelineDB.State = 0 Then
		objConnpipelineDB.Close()
	End If
'objConnpipelineDB  = nothing
End sub


'********************************************************************************************************************************
Function GetConnpipelineDB_AUDIO
	If objConnpipelineDB_AUDIO.State = 0 Then
		objConnpipelineDB_AUDIO.ConnectionTimeout = Application("audioDB_ConnectionTimeout")
		objConnpipelineDB_AUDIO.CommandTimeout = Application("audioDB_CommandTimeout")
		objConnpipelineDB_AUDIO.CursorLocation = Application("audioDB_CursorLocation")
		objConnpipelineDB_AUDIO.Open Application("audioDB_ConnectionString"), Application("audioDB_RuntimeUserName"), Application("audioDB_RuntimePassword")
	End If
	Set GetConnpipelineDB_AUDIO = objConnpipelineDB_AUDIO
End Function

sub CloseConnpipelineDB_AUDIO
	If objConnpipelineDB_AUDIO.State = 0 Then
		objConnpipelineDB_AUDIO.Close()
	End If
'objConnpipelineDB_AUDIO  = nothing
End sub
'********************************************************************************************************************************


'p1 es fecha en afiliacion 
'p2 es hora
'p3 es consecutivo
'p4 es unidad cuando haya 0 cuando no
Sub writelog3(pagina, ctr, producto,plan,error,code,sp,texto)
dim Sql
dim usuario
dim nrodocum
dim p1,p2,p3,p4
'on error resume next

if Session.Contents("sp_userId") <> "" then
	usuario = Session.Contents("sp_userId")  
else
	usuario = 0	
end if
if Session.Contents("sp_docNumber") <> "" then
	nrodocum = Session.Contents("sp_docNumber")  
else
	nrodocum = 0
end if
p1 = 0
p2 = 0
p3 = 0
p4 = 0
Sql = "exec sppl_putlog " + cstr(Session.sessionID)+","+cstr(pagina)+",'"+ _
Request.ServerVariables("REMOTE_ADDR")+"','"+sp+"',"+cstr(usuario)+",'"+ _ 
cstr(ctr)+"','"+producto+"',"+cstr(nrodocum)+","+cstr(error)+",'"+plan+"'" +  _
",'"+texto+"',"+ _
cstr(p1)+","+cstr(p2)+","+cstr(p3)+","+cstr(p4)
set rs=server.CreateObject("ADODB.Recordset")
'Abre la conexion 
set cn  = GetConnpipelineDB
'Ejecuta el stored procedure log2
cn.execute Sql
exit sub
errLog:
Response.Redirect "../error/critical_error_message.asp?mensaje=Ha ocurrido una visicitud, grabando el evento anterior. Por favor reintente en un momento o de aviso a nuestro call center."
end sub

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Error trapping functions added by Andres F. Orozco
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getTextStream(filename, iomode, create, AsciiOrUnicode)
	Dim fs, Path
	const forReading=1
	const forWriting=2
	const forAppending=8
	const TristateFalse=0
	const TristateTrue=-1
	const TristateUseDefault=-2
	const TristateMixed=-2
	Set fs=server.createObject("Scripting.FilesystemObject")
	If isEmpty(iomode) then Iomode=ForReading
	If isEmpty(create) then Create=False
	If isEmpty(AsciiOrUnicode) then AsciiOrUnicode= TriStateUseDefault
	Set getTextStream=fs.openTextFile(filename, iomode, create, AsciiOrUnicode)
End Function

Sub logError(errorFile, ErrNumber, ErrSource, ErrDescription, SourcePage, UserLogin, UserName)
  	Dim ts
	Set ts = getTextStream(errorFile, 8, True,Empty)
	ts.writeLine formatDateTime(now) & "; " & ErrNumber & "; " & ErrSource & "; " & ErrDescription & _
	"; " & SourcePage & "; " & UserLogin & "; " & UserName & "; " & Session.SessionID
	ts.close
	set ts = nothing
End Sub

'=====================================================================
'This function checks if the web site is available
'=====================================================================
Function Available()
	Dim objConn, objRst, strSQL
	Set objConn = GetConnpipelineDB
	Set objRst = Server.CreateObject("ADODB.Recordset")
	strSQL = "spsp_GetDBStatus"
	objRst.Open strSQL, objConn, 3
	If Not(objRst.BOF And objRst.EOF) Then
		If objRst.Fields(0) = "N" Then
			Available = 1 'Not available
		Elseif objRst.Fields(1) = "N" Then
			Available = 2 'Partially available
		Else
			Available = 0 'Available
		End If
	Else
		Available = 0
	End If
	Set objRst = Nothing
	CloseConnpipelineDB
	Set objConn = Nothing
End Function



'==========================================================================
'<I&T - MENCO 2012/06/08 - Proyecto Capital+Seguro>
' Realiza la invocaci?n del web service TaxFacade para obtener la informaci?n del saldo (incluyendo comisi?n)
'==========================================================================
Function InvokeTaxFacade(url)
    Dim xmlDOC
    Dim bOK
    Dim HTTP
    Dim accion

    Set HTTP = CreateObject("MSXML2.XMLHTTP")
    Set xmlDOC = CreateObject("MSXML.DOMDocument")

    xmlDOC.Async = False
    accion = Application.Contents("UrlTaxFacade") & "/" & url


    HTTP.Open "GET", accion, false
    HTTP.Send(null)


    bOK = xmlDOC.load(HTTP.responseXML)
    If Not bOK Then
      Response.Write HTTP.responseXML
      Exit Function
    End If
    Dim Result
    Result = ""
    Result = xmlDOC.Text
    InvokeTaxFacade = Result
End Function

</SCRIPT>
