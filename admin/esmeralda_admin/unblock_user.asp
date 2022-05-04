<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../../operations/_pipeline_scripts/mailsender.asp"-->
<%
'===================================================================================
'File Name:		unblock_user.asp 8503
'Path:				esmeralda_admin/
'Created By:		A. Orozco 2001/06/13
'Last Modified:	A. Orozco 2001/09/19
'						A. Orozco 2001/10/11
'Parameters:	
'Returns:		
'Additional Information:
'				Adiminstration page.
'				Adapted to manacha
'===================================================================================

'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1

Dim strPass, arrChars, I,J
Dim adoConn, rstNewMail, rstLog, arrUser, arrLogin, arrMail
Dim Sql, status
Dim pcmd
Dim strOpening, strNames, strNewSA, strSubject, rsMessages, Body, cnPip
Dim processInfo, component_id, valueNewLog

Authorize 5,11
arrUser = Split(Request.QueryString("user"),", ")
arrChars = Split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,a,b,c,d,e,f,g,h,i," & _
"j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z",",")


Set cnPip = GetConnPipelineDB

' Set rstNewMail = Server.CreateObject("ADODB.Recordset")

%>
<html>
	<head>
		<meta http-equiv="X-UA-Compatible" content="IE=9" />
		<link href="../../css/style.css" rel="stylesheet" type="text/css"/>
		<link href="../../css/CompiledGlobal.css" rel="stylesheet" type="text/css"/>
	</head>
	<body class="cuerpo">
		<div>
			<h1>
				Nueva contrase�a y desbloqueo de usuario
			</h1>
			<table width="90%" style="margin:0 auto;">
				<tr>
					<td width="11%">
					<h2>Generaci�n de contrase�as</h2>
					</td>
				</tr>			
<%
For I = 0 To UBound(arrUser)
	Dim user, mail, Nombres, Apellidos
	' Sql = "spem_GetDatosUsuarioXuserID '" & arrUser(I) & "'"
	' rstNewMail.Open Sql,cnPip

	Set pcmd = Server.CreateObject("ADODB.Command")
	pcmd.CommandText = "spem_GetDatosUsuarioXuserID"
	pcmd.CommandType = 4 'adCmdStoredProc
	pcmd.ActiveConnection = cnPip

	pcmd.Parameters.Append(pcmd.CreateParameter("@user", 200, 1, 10, arrUser(I))) '200=adVarChar, 1=adInput

	Set rstNewMail = pcmd.Execute

	user = rstNewMail.Fields("login")
	mail = trim(rstNewMail.Fields("email"))
	Nombres = rstNewMail.Fields("Nombres")
	Apellidos = rstNewMail.Fields("Apellidos")
	rstNewMail.Close

	component_id = "unblock_user.asp"
	processInfo = "admin/esmeralada_admin/unblock_user.asp " & "Loaded by " & Session("sp_miLogin")


	write_sp_log cnPip, 8503, "spem_GetDatosUsuarioxuserID", 0, "", "", 0, 0, "", "admin/esmeralada_admin/unblock_user.asp " & _
	"Loaded by " & Session("sp_miLogin")

	' Sql = "spsp_BlockUser '" + user + "','A'"
	' cnPip.Execute Sql

	Set pcmd = Server.CreateObject("ADODB.Command")
	pcmd.CommandText = "spsp_BlockUser"
	pcmd.CommandType = 4 'adCmdStoredProc
	pcmd.ActiveConnection = cnPip

	pcmd.Parameters.Append(pcmd.CreateParameter("@login", 200, 1, 25, user)) '200=adVarChar, 1=adInput
	pcmd.Parameters.Append(pcmd.CreateParameter("@state", 200, 1, 1, "A")) '200=adVarChar, 1=adInput

	pcmd.Execute

	write_sp_log cnPip, 8503, "spsp_BlockUser", 0, "", "", 0, 0, "", "admin/esmeralada_admin/unblock_user.asp " & _
	"Loaded by " & Session("sp_miLogin")
	Randomize
	strPass = Int(10000 * Rnd)
	For J = 0 to 1 
		strPass = strPass & arrChars(Int((UBound(arrChars) + 1) * Rnd))
	Next
	' Sql = "spsp_SetNewPwd '" & user & "','" & strPass & "'"
	' rstNewMail.Open Sql, cnPip

	Set pcmd = Server.CreateObject("ADODB.Command")
	pcmd.CommandText = "spsp_SetNewPwd"
	pcmd.CommandType = 4 'adCmdStoredProc
	pcmd.ActiveConnection = cnPip

	pcmd.Parameters.Append(pcmd.CreateParameter("@login", 200, 1, 25, user)) '200=adVarChar, 1=adInput
	pcmd.Parameters.Append(pcmd.CreateParameter("@pwd", 200, 1, 40, strPass)) '200=adVarChar, 1=adInput

	Set rstNewMail = pcmd.Execute

	Set rsMessages = Server.CreateObject("ADODB.Recordset")
	Sql = "spem_GetEmailText 'NewPin','S'"
	rsMessages.Open Sql,cnPip
	strSubject = rsMessages.Fields("texto")
	rsMessages.Close
	
	strOpening = "Apreciado(a) Cliente:"
	
	Sql = "spem_GetEmailText 'NewPin','H0'"
	rsMessages.Open Sql,cnPip
	Body = rsMessages.Fields("texto") & strOpening
	rsMessages.Close
	Sql = "spem_GetEmailText 'NewPin','H1'"
	rsMessages.Open Sql,cnPip
	'Body = Body & rsMessages.Fields("texto") & strPass
	
	valueNewLog = "Usuario:"&user&",Nueva contrasena: ********"

	Body = "<h3>Cliente:</h3><p>" & Nombres & " " & Apellidos & "</p><h3>Usuario:</h3><p>"&  user & "</p><h3>Nueva contrase�a:</h3><span style='font-size:15px;'>" & strPass & "</span>" 
	rsMessages.Close
	Sql = "spem_GetEmailText 'NewPin','B0'"
	rsMessages.Open Sql,cnPip
	'Body = Body & rsMessages.Fields("texto")
	rsMessages.Close
	Sql = "spem_GetEmailText 'NewPin','F0'"
	rsMessages.Open Sql,cnPip
	'Body = Body & rsMessages.Fields("texto")
	rsMessages.Close
	Set rsMessages = Nothing

	'SendMail Application("SenderMail"),mail,strSubject,Body
	'Response.Write "<p>"
	%>
				<tr>
					<td>
						<% Response.Write Body %>
					</td>
				</tr>
	<%
	write_sp_log cnPip,8503, "Cambio pwd y desbloquear usuario: " & mail, 0, "", "", 0, 0, "", "admin/esmeralada_admin/unblock_user.asp " & _
	"Loaded by " & Session("sp_miLogin")

	write_dataLog Response.Status,component_id,processInfo,Session.contents("idworker"), "spem_GetDatosUsuarioXuserID" ,"",valueNewLog,"Operación-Modificación","N/A"

Next

CloseConnPipelineDB
Set cnPip = Nothing
Set rstNewMail = Nothing
Set rstLog = Nothing
%>
			</table>
			<br>
			<br>
			<p style="text-align:center;">
				Los usuarios seleccionados han sido desbloqueados y se les ha asignado una nueva contrase�a. Recuerde que la contrase�a es sensible a may�sculas
			</p>
			<br />
			<br />
			<center>
				<input name="return" onclick="javascript:window.location='admin_menu.asp'" type="button" value="Retornar"/>
			</center>
		</div>
	</body>
</html>
<%
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>