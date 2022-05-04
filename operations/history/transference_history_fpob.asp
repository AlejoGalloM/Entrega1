<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		transaction_history.asp 11300
'Path:			transaction_history/
'Created By:	Andres Felipe Orozco 2001/05/31
'Last Modified:	Andres Felipe Orozco 2001/06/06
'Last Modified:	Guillermo Aristizabal  2001/09/21
'						A. Orozco 2001/10/08
'						Guillermo Aristizabal 2001/10/11
'Parameters:	User must be logged on
'				Contract, Product
'Returns:		Transaction history for the selected contract-product
'Additional Information:
'===================================================================================
Option Explicit
On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../../_ScriptLibrary/pm.asp"-->
<!--#include file="../_pipeline_scripts/url_check.asp"-->
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../../operations/transactions/reglamentoscriptsPipeline.asp"-->
<%

Authorize 1,15
Dim rs, cn 'Recordset, Connection
Dim rsHistory
Dim Sql
Dim adoConn
Dim I,J 'Recordset counters
Dim arrHistory
Dim ProdName
Dim Contract, Product, Plan, ClientId
dim Reference  '<I&T - DMPC>
Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
ClientId = Request.Form("ClientId")

'==========================================================================
''<I&T - DMPC 2009/03/31 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia única a partir del Contrato y el producto
'==========================================================================
 Reference = GetReferenciaUnica( Product, Contract)
'==========================================================================

Set cn  = GetConnpipelineDB
Set rsHistory = Server.CreateObject("ADODB.Recordset")
Sql = "Sppl_GetTrasladosFPOB " & ClientId
rsHistory.Open Sql,cn
If rsHistory.BOF And rsHistory.EOF Then
	arrHistory = 0
Else
	arrHistory = rsHistory.GetRows()
End If
rsHistory.Close
Set rsHistory = Nothing
write_sp_log cn, 11300, "Sppl_GetTrasladosFPOB", Contract, Product, Plan, ClientId, 0, "", "transference_history_fpob.asp " & _
"Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"transference_history_fpob.asp","transference_history_fpob.asp " &"- " & Session("sp_miLogin"),Session.contents("name"),Sql ,"N/A","null","Consulta","N/A"

CloseConnpipelineDB

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src='../_pipeline_scripts/validation.js'></SCRIPT>
<%
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	OpenBody "", ""
OpenTable "90%","'' align=center"
	OpenTr ""
		OpenTd "",""
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Traslados<br>"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "tbody","colspan=2"
						Response.Write "Producto :"
					CloseTd
					OpenTd "tbody","colspan=3 width=83%"
						Response.Write Product
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "tbody","colspan=2"
						'Response.Write "Afiliación :" '<I&T - DMPC - Modificado por Proceso de Referencia Unica - Recaudos>
						Response.Write "Contrato :" 
					CloseTd
					OpenTd "tbody","colspan=3 width=83%"
						'Response.Write Contract
						'<I&T - DMPC 2009/03/31 - Modificado por proceso de Referencia Única - Recaudos>
						if IsNull(Reference) or len(Reference)>12 Then
							Response.Write Contract
						else
							Response.Write Reference
						end if
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "thead","colspan=5 align=left"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			CloseTable
			'Check if there's transference history
			If IsArray(arrHistory) Then
				OpenTable "100%","'' align=center border=1"
					OpenTr "class=teven align=center"
						OpenTd "thead","width=12%"
							Response.Write "Tipo Traslado"
						CloseTd
						OpenTd "thead","width=10%"
							Response.Write "Entidad"
						CloseTd
						OpenTd "thead","width=12%"
							Response.Write "Estado Traslado"
						CloseTd
						OpenTd "thead","width=11%"
							Response.Write "Estado Aprobación"
						CloseTd
						OpenTd "thead","width=11%"
							Response.Write "Fecha Solicitud"
						CloseTd
						OpenTd "thead","width=10%"
							Response.Write "Fecha Respuesta"
						CloseTd
						OpenTd "thead","width=13%"
							Response.Write "Fecha Pago"
						CloseTd
						OpenTd "thead","width=13%"
							Response.Write "Fecha Legalización"
						CloseTd
						OpenTd "thead","width=8%"
							Response.Write "Valor"
						CloseTd
					CloseTr
					For J = 0 To UBound(arrHistory,2)
						If (J Mod 2) = 0 Then
							OpenTr "class=todd align=center"
						Else
							OpenTr "class=teven align=center"	
						End If
						For I = 0 To UBound(arrHistory)
							Select Case I
								Case 2,5,7,8
									OpenTd "tbody",""
										Response.Write arrHistory(I,J)
									CloseTd
								Case 9,10,11,12
									OpenTd "tbody",""
										If Not(IsNull(arrHistory(I,J))) Then
											Response.Write FormatDateTime(arrHistory(I,J),2)
										Else
											Response.Write "&nbsp;"
										End If
									CloseTd
								Case 13
									OpenTd "tbody",""
										Response.Write FormatCurrency(arrHistory(I,J),2)
									CloseTd
							End Select
						Next
						CloseTr
					Next
				CloseTable
			Else
				OpenTable "100%","'' align=center border=1"
					OpenTr "class=teven align=center"
						OpenTd "thead","width=12%"
							Response.Write "No hay historia de traslados"
						CloseTd
					CloseTr
				CloseTable
			End If
		CloseTd
	CloseTr
CloseTable
CloseBody
CloseHTML
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>