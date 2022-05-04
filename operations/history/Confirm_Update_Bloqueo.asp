<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:					Confirm_Debits.asp 0021
'Path:							History/
'Created By:					Rafael Lagos 2002/10/17
'Parameters:					
'Returns:						
'Additional Information:	
'===================================================================================
Option Explicit
On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../../operations/transactions/reglamentoscriptsPipeline.asp"-->
<%
Authorize 0,24
Dim objConn 'Database Connection
Dim objRst,objUser,rsUser,rsUser1 'Recordset object
'Dim objRst_Pend
Dim strSQL, Sql 'Query container
Dim Sel, I, J, Contrato_Text, Observacion
Dim arrAff,ValUser
Dim Count, Id1, Id2, Id3
Dim S_Date, E_Date, Unit, Tipo, Confirmation

dim arrReference  ' <I&T - DMPC>
dim mensaje  ' <I&T - DMPC>
dim Contrato ' <I&T - DMPC>
dim Producto ' <I&T - DMPC>
Dim processInfo, component_id, valueLog

Set objConn = GetConnPipelineDB

write_sp_log objConn, 002100, "", 0, "", "", 0, 0, "", "Confirm_Update_Bloqueo.asp loaded by " & _
Session.Contents("sp_milogin")


Observacion = Request.Form("txtObservacion")
Tipo = Request.QueryString("Tipo")
Contrato_Text = Request.Form("txtContrato")


Set objConn = GetConnPipelineDB
Sql =  "Upd_BloqueoAportes '" & Contrato_Text  & "','" &  Tipo & "','" & Observacion & "','"& Session.Contents("sp_milogin") &"'"
'response.write "SQL = " & Sql 
Set rsUser1 = Server.CreateObject("ADODB.Recordset")
rsUser1.Open Sql, objConn 

write_sp_log objConn, 15400, "Upd_BloqueoAportes", 0, "", "", 0, 0, "", "Confirm_Update_Bloqueo.asp loaded by " & _
Session.Contents("sp_milogin")

component_id = "Confirm_Update_Bloqueo.asp"
processInfo =  "Confirm_Update_Bloqueo.asp loaded by " & Session.Contents("sp_milogin")
valueLog = "Observacion: "&Observacion&", Tipo: "&Tipo&", Texto Contrato: "&Contrato_Text

write_dataLog Response.Status,component_id,processInfo,Session.contents("idworker"), "Upd_BloqueoAportes" ,"",valueLog,"Operación-Modificación","N/A"


'==========================================================================
''<I&T - DMPC 2009/02/11 - Modificado por proceso de Referencia Unica>
' Valida el n�mero de contrato digitado, revisando el d�gito de verificaci�n
'==========================================================================
arrReference = GetReferenciaUnicaInversa (Contrato_Text, "null")

 mensaje=arrReference(2,0)

If Mensaje<>"null" Then
			Response.Write("<SCRIPT LANGUAGE='javascript'>") 
			Response.Write("alert('" + mensaje+ " Recuerde que el n�mero del contrato est� compuesto por 12 d�gitos, verif�quelo y vuelva a intentar."+ "');") 
			Response.Write("window.location.href = 'BloqueoAportes_Update.asp';")
			Response.Write("</SCRIPT>")    
	else
			Contrato=arrReference(0,0)
		    Producto=arrReference(1,0)
	End If 
'==========================================================================

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src=../operations/_pipeline_scripts/validation.js></SCRIPT>
<%
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	OpenBody "", ""
	
		OpenTable "75%","'' align=center"
			OpenTr ""
				OpenTd "",""
					Response.Write "&nbsp;"
				CloseTd
			CloseTr
			OpenTr "class=teven"
				OpenTd "thead", "align=center colspan=2"
					Response.Write "Contrato Desbloqueado"
				CloseTd
			CloseTr
		CloseTable
	
		OpenTable "75%", "'' border = 1 align=center"
			OpenTr "class=tbody"
				OpenTd "","width=40"
					Response.Write "Contrato :"
				CloseTd
				OpenTd "","width=60"
					 Response.Write Contrato_Text 
				CloseTd
			CloseTr
			OpenTr "class=tbody"
				OpenTd "tbody", ""
					Response.Write "Descripcion :"
				CloseTd
				OpenTd "tbody", ""
					Response.Write Observacion
				CloseTd
			CloseTr
			
					
					
		CloseTable
	
		CloseBody
CloseHTML

CloseConnPipelineDB


If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
