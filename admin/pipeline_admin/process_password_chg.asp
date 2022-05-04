<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		process_password_chg.asp.asp 6303
'Path:				admin/pipeline_admin/
'Created By:		Guillermo Aristizabal 2001/08/02
'Last Modified:	A. Orozco 2001/09/17
'						A. Orozco 2001/10/11	Changed page_id=13363 and added spsp_PwdChange in write_sp_log
'			Diana Mariced P�rez	2008/05/08	
'			I&T  - Oscar Diaz 	2010/08/27
'Additional Information:
'===================================================================================
Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<%

Function ManejarErrorStr(strError)
	ManejarErrorStr = Mid(strError , 48 , Len(strError))
End function

Dim sql
Dim cn, rs, rslog
Dim adoConn
Dim sqllog
Dim login
Dim currentPassword
Dim newPassword
Dim newPassword2
Dim changePassword_ok
Dim attempts
Dim numDigits
Dim Body
'I&T - Oscar Diaz -08/25/2010 Variables de parametros de configuraci�n
dim sqlHisto
dim sqlParam
dim sqlActualHisto
dim rsHisto
dim rsParam
dim rsActualHisto
dim regularExpresion
dim vigenciaMax
dim vigenciaMin
dim historicoClaves
dim regularExpRepeat
dim arrParam
dim maxLength
dim minLength
dim cantidadHisto
dim maxAttempts
dim pcmd
Dim processInfo, component_id
'I&T - Oscar Diaz  -08/25/2010 


login = Session("sp_miLogin")
currentPassword = Request.Form("txtOldPass")
newPassword = Request.Form("txtNewPass")
newPassword2 = Request.Form("txtNewPass2")

component_id = "process_password_change.asp"
processInfo =  "process_password_change.asp " & "- " & Session("sp_miLogin")

write_dataLog Response.Status,component_id,processInfo,Session.contents("idworker"), "adminsmlv.asp" ,currentPassword,"********","Operación-Modificación","N/A"

'I&T  Se obtiene el historico de contrase�as del usuario
'Adem�s se obtiene los parametros de configuraci�n de la clave
set cn = GetConnpipelineDB

'set rsParam = Server.CreateObject("ADODB.RecordSet") 
'sqlParam = "usuarios..Get_ParametroClave " & Application("SiteParam")
'
'rsParam.Open sqlParam,cn

Set pcmd = Server.CreateObject("ADODB.Command")
pcmd.CommandText = "usuarios..Get_ParametroClave"
pcmd.CommandType = 4 'adCmdStoredProc
pcmd.ActiveConnection = cn
pcmd.Parameters.Append(pcmd.CreateParameter("@sitio", 3, 1, , Application("SiteParam"))) '3=adInteger, 1=adInput
Set rsParam = pcmd.Execute

'Validaciones preliminares
If rsParam.BOF And rsParam.EOF Then
	arrParam = 0
Else
	arrParam = rsParam.GetRows()
End If
rsParam.Close 
	
If IsArray(arrParam) Then
	regularExpresion = arrParam(9,0)
	vigenciaMax = arrParam(0,0)
	vigenciaMin = arrParam(1,0)
	historicoClaves = arrParam(7,0)
	regularExpRepeat = arrParam(10,0)
	maxLength = arrParam(11,0)
	minLength = arrParam(2,0)
	maxAttempts = arrParam(16,0)
End if

If Err.number <> 0 Then
	CloseConnpipelineDB
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If


'I&T - Oscar Diaz - 25/08/2010
'Se obtiene el historico de contrase�as del usuario para hacer las respectivas validaciones
' set rsHisto=server.CreateObject("ADODB.Recordset")
' sqlHisto = "EXEC usuarios..Get_HistoricoClave_Clave '"+ login + "','"+ newPassword +"','" + currentPassword + "'"
' rsHisto.Open sqlHisto, cn

Set pcmd = Server.CreateObject("ADODB.Command")
pcmd.CommandText = "usuarios..Get_HistoricoClave_Clave"
pcmd.CommandType = 4 'adCmdStoredProc
pcmd.ActiveConnection = cn

pcmd.Parameters.Append(pcmd.CreateParameter("@login", 200, 1, 25, login)) '200=adVarChar, 1=adInput
pcmd.Parameters.Append(pcmd.CreateParameter("@pwd", 200, 1, 40, newPassword)) '200=adVarChar, 1=adInput
pcmd.Parameters.Append(pcmd.CreateParameter("@oldpwd", 200, 1, 40, currentPassword)) '200=adVarChar, 1=adInput

Set rsHisto = pcmd.Execute

If Err.number <> 0 Then
	CloseConnpipelineDB
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If

if(rsHisto(0) = 0) Then	'Historico
	CloseConnpipelineDB
	
	'I&T - Oscar Diaz
	'Se obtiene el historico de contrase�as del usuario para hacer las respectivas validaciones
		
	cn = GetConnpipelineDB
	' set rsActualHisto=server.CreateObject("ADODB.Recordset")
	' sqlActualHisto = "EXEC usuarios..Get_HistoricoClave_Actual '"+ login + "','" + currentPassword + "'"
	' rsActualHisto.Open sqlActualHisto, cn

	Set pcmd = Server.CreateObject("ADODB.Command")
	pcmd.CommandText = "usuarios..Get_HistoricoClave_Actual"
	pcmd.CommandType = 4 'adCmdStoredProc
	pcmd.ActiveConnection = cn

	pcmd.Parameters.Append(pcmd.CreateParameter("@login", 200, 1, 25, login)) '200=adVarChar, 1=adInput
	pcmd.Parameters.Append(pcmd.CreateParameter("@oldpwd", 200, 1, 40, currentPassword)) '200=adVarChar, 1=adInput

	Set rsActualHisto = pcmd.Execute

	cantidadHisto = rsActualHisto(0)

	response.write " despues de suarios..Get_HistoricoClave_Actual "
	response.end
	
	If Err.number <> 0 Then
		Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
		"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
		Server.URLEncode(Err.source))
	End If
	
	if(cantidadHisto < vigenciaMin )then 'Cantidad que puede cambiar contrase�a en el d�a
		CloseConnpipelineDB
		
		Set cn = GetConnpipelineDB
		' set rs = server.CreateObject("ADODB.Recordset")

		' sql = "exec usuarios..sp_cambiarClave '" + login + "','"+ newPassword +"','" + currentPassword + "'"

		' rs.Open sql,cn

		Set pcmd = Server.CreateObject("ADODB.Command")
		pcmd.CommandText = "usuarios..sp_cambiarClave"
		pcmd.CommandType = 4 'adCmdStoredProc
		pcmd.ActiveConnection = cn

		pcmd.Parameters.Append(pcmd.CreateParameter("@login", 200, 1, 25, login)) '200=adVarChar, 1=adInput
		pcmd.Parameters.Append(pcmd.CreateParameter("@pwd", 200, 1, 40, newPassword)) '200=adVarChar, 1=adInput
		pcmd.Parameters.Append(pcmd.CreateParameter("@oldpwd", 200, 1, 40, currentPassword)) '200=adVarChar, 1=adInput

		Set rs = pcmd.Execute

		' 21/Ago/2009 - I&T - Armando J. Arias G�mez 
		' Valida los errores que se pueden generar en el sp_cambiarClave por el trigger tr_usuario
		if Err.number <> 0 then
				Session.Contents("sp_error")=ManejarErrorStr(Err.Description)
				CloseConnpipelineDB
				Response.Redirect "password_change.asp" 
		end if

		changePassword_ok = rs(0)
		attempts = rs(1)

		'write_sp_log cn, 6303, "", 0, "", "", 0, 0, "", "process_password_change.asp " & _
		write_sp_log cn, 13363, "usuarios..sp_cambiarClave", 0, "", "", 0, 0, "", "process_password_change.asp " & _
		"- " & Session("sp_miLogin")

		if changePassword_ok = 1 then
			Session("pwdchg") = 1
			Session.Contents("sp_error") = ""
			CloseConnpipelineDB
			If Err.number <> 0 Then
				Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
				"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
				Server.URLEncode(Err.source))
			End If
		%>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.location = '../<%=Session("sp_init_page")%>'
			//-->
			</SCRIPT>
		<%
		elseif changePassword_ok = 2 then
			if (attempts < maxAttempts) then
				Session.Contents("sp_error")=Application("error0")
				CloseConnpipelineDB
				If Err.number <> 0 Then
					Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
					"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
					Server.URLEncode(Err.source))
				End If			
				Response.Redirect "password_change.asp" 
			else
				rs.Close
				' sql = "spsp_BlockUser '" + login + "', 'C'"				
				' rs.Open sql,cn
			
				Set pcmd = Server.CreateObject("ADODB.Command")
				pcmd.CommandText = "spsp_BlockUser"
				pcmd.CommandType = 4 'adCmdStoredProc
				pcmd.ActiveConnection = cn

				pcmd.Parameters.Append(pcmd.CreateParameter("@login", 200, 1, 25, login)) '200=adVarChar, 1=adInput
				pcmd.Parameters.Append(pcmd.CreateParameter("@state", 200, 1, 1, "C")) '200=adVarChar, 1=adInput

				Set rs = pcmd.Execute

				Dim strOpening, strNames, strNewSA, rsMessages, strSubject
				Set rsMessages = Server.CreateObject("ADODB.Recordset")
				Sql = "spem_GetEmailText 'LockedPin','S'"
				rsMessages.Open Sql, cn
				strSubject = rsMessages.Fields("texto")
				rsMessages.Close
						
				Select Case Session.Contents("sex") 
					Case "M"
						strOpening = "Apreciado "
					Case "F"
						strOpening = "Apreciada "
					Case Else
						strOpening = "Apreciado(a) "
				End Select
				If Session.Contents("metaname") = "" Then
					strNames = "Cliente :"
				Else
					strNames = Session.Contents("metaname")
				End If

				Sql = "spem_GetEmailText 'LockedPin','H0'"
				rsMessages.Open Sql, cn
				Body = rsMessages.Fields("texto") & " " & strOpening & " " & StrNames
				rsMessages.Close
						
				Sql = "spem_GetEmailText 'LockedPin','B0'"
				rsMessages.Open Sql, cn
				Body = Body & rsMessages.Fields("texto") 
				rsMessages.Close

				Sql = "spem_GetEmailText 'LockedPin','M1'"
				rsMessages.Open Sql, cn
				Body = Body & rsMessages.Fields("texto") 
				rsMessages.Close


				Sql = "spem_GetEmailText 'LockedPin','B1'"
				rsMessages.Open Sql, cn
				Body = Body & rsMessages.Fields("texto") 
				rsMessages.Close

				Sql = "spem_GetEmailText 'LockedPin','F0'"
				rsMessages.Open Sql, cn
				Body = Body & rsMessages.Fields("texto") 
				rsMessages.Close

				CloseConnpipelineDB
				If Err.number <> 0 Then
					Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
					"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
					Server.URLEncode(Err.source))
				End If
				%>
				<SCRIPT LANGUAGE=javascript>
				<!--
					window.parent.location = '../../login/sorry_page.asp?error=1'
				//-->
				</SCRIPT>
				<%
			end if			
		end if
	else
		Session.Contents("sp_error")=Application("error7") 
		If Err.number <> 0 Then
			Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
			"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
			Server.URLEncode(Err.source))
		End If			
		Response.Redirect "password_change.asp" 
	end if 'Historico Actual
else
	Session.Contents("sp_error")=Application("error6")
	
	response.write " despues de sp_error" & Application("error6")
	response.end

	
	If Err.number <> 0 Then
		Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
		"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
		Server.URLEncode(Err.source))
	End If			
	Response.Redirect "password_change.asp" 
end if 'Historico		
CloseConnpipelineDB
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>