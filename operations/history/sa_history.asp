<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		sa_history.asp 702
'Path:				history/
'Created By:		A. Orozco 2001/10/25
'Last Modified:	A. Orozco 2001/11/01
'Parameters:		User must be logged on
'						Contract, Product, Plan
'Returns:			Standing Allocation history for the selected contract-product
'Additional Information:
'===================================================================================
Option Explicit
'On Error Resume Next
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

Authorize 7,1

Dim rs, cn 'Recordset, Connection
Dim rsHistory
Dim Sql
Dim adoConn
Dim I,J 'Recordset counters
Dim arrHistory
Dim Contract, Product, Plan, Name
Dim EffDate
Dim Sel
Dim RecCount, PgCount, Pages
dim Reference '<I&T - DMPC>
Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")

'==========================================================================
''<I&T - DMPC 2009/03/30 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia única a partir del Contrato y el producto
'==========================================================================
 dim ContratoFCO

	If Product <> "FCO" Then
		Reference = GetReferenciaUnica( Product, Contract)
	else
		ContratoFCO = Plan + Trim(Contract)
		Reference = GetReferenciaUnica( Product, ContratoFCO)
	end if
'==========================================================================

'Get transactions history 
Sql = "spsp_GetSAHistory " & Contract

Set rsHistory = Server.CreateObject("ADODB.Recordset")
Set cn  = GetConnpipelineDB

write_sp_log cn, 702, "spsp_GetSAHistory", Contract, Product, Plan, 0, 0, "", "sa_history.asp loaded " & _
"by " & Session("sp_miLogin")
rsHistory.PageSize = Application("PagesHistory")

rsHistory.Open Sql,cn
Pages = rsHistory.PageCount
CloseConnpipelineDB
Set cn = Nothing

write_dataLog Response.Status,"sa_history.asp","spsp_GetSAHistory " &"- " & Session("sp_miLogin"),Session.contents("name"),Sql ,"N/A","null","Consulta","N/A"

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src='../_pipeline_scripts/validation.js'></SCRIPT>
<%
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	OpenBody "''", "bgcolor='#FFFFFF' text='#000000'"
OpenTable "90%", "'' align=center"
	OpenTr ""
		OpenTd "", ""
			OpenTable "100%", "'' align=center"
				OpenTr ""
					OpenTd "thead", "align=left"
						Response.Write "<h3>Historia Standing Allocation</h3>"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead", "align=left"
						Response.Write "Este reporte histórico únicamente despliega las modificaciones de Standing Allocation efectuadas por " & _
						"Pipeline o Portal de Clientes. Todas la modificaciones previas efectuadas directamente en AS400 no son visualizadas"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%", "'' align=center"
				OpenTr ""
					OpenTd "tbody", "align=left"
						Response.Write "Producto :"
					CloseTd
					OpenTd "tbody", "align=left"
						'Get Product Description
						Response.Write Product
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "tbody", "align=left"
						'Response.Write "Afiliación :" '<I&T - DMPC 2009/03/30 - Modificado por proceso de Referencia Única - Recaudos>
						Response.Write "Contrato :"
					CloseTd
					OpenTd "tbody", "align=left"
						'<I&T - DMPC 2009/03/30 - Modificado por proceso de Referencia Única - Recaudos>
						'Response.Write Contract
						if IsNull(Reference) or len(Reference)>12 Then
							Response.Write Contract		
						else
							Response.Write Reference
						end if	
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "tbody", "align=left"
						Response.Write "&nbsp;"
					CloseTd
					OpenTd "tbody", "align=left"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%", "'' align=center"
				OpenTr ""
					OpenTd "thead", ""
					CloseTd
				CloseTr
			CloseTable
			If rsHistory.BOF And rsHistory.EOF Then
				OpenTable "100%", "'' border=1 align=center"
					OpenTr "class=teven"
						OpenTd "thead", "align=center"
							Response.Write "No hay historia de standing allocation"
						CloseTd
					CloseTr
				CloseTable
			Else
				If Request.Form("Page") = "" Then
					PgCount = 1
				Else
					PgCount = Request.Form("Page")
				End If
				rsHistory.AbsolutePage = PgCount
				OpenTable "100%", "'' border=1 align=center"
					OpenTr ""
						OpenForm "pages", "post", "", ""
						OpenTd "tbody", "align=center colspan=7"
							For I = 1 To Pages
								PlaceInput "page" & I, "button", I, "class=sbttn onClick='javascript:pages.Page.value=" & _
								I & "; pages.submit()'"
							Next
						CloseTd
							PlaceInput "Page", "hidden", "", ""
							PlaceInput "Contract", "hidden", Contract, ""
							PlaceInput "Product", "hidden", Product, ""
							PlaceInput "Plan", "hidden", Plan, ""
						CloseForm
					CloseTr
					OpenTr ""
						OpenForm "prev", "post", "", ""
						OpenTd "thead", "align=center nowrap"
							If  PgCount > 1 Then
									PlaceInput "submit","submit", "< Anterior", "class=sbttn"
									PlaceInput "Page", "hidden", PgCount - 1, ""
									PlaceInput "Contract", "hidden", Contract, ""
									PlaceInput "Product", "hidden", Product, ""
									PlaceInput "Plan", "hidden", Plan, ""
							Else
								Response.Write "&nbsp;"
							End If
						CloseTd
						CloseForm
						OpenTd "thead", "colspan=3 align=center nowrap"
							Response.Write "Pagina: " & PgCount & " de " & Pages
							Response.Write " - No. Total de registros: " & rsHistory.RecordCount
						CloseTd
						OpenForm "next", "post", "", ""
						OpenTd "thead", "align=center nowrap"
							If  rsHistory.AbsolutePage < Pages Then
									PlaceInput "submit","submit", "Siguiente >", "class=sbttn"
									PlaceInput "Page", "hidden", PgCount + 1, ""
									PlaceInput "Contract", "hidden", Contract, ""
									PlaceInput "Product", "hidden", Product, ""
									PlaceInput "Plan", "hidden", Plan, ""
							Else
								Response.Write "&nbsp;"
							End If
						CloseTd
						CloseForm
					CloseTr
					OpenTr "class=teven"
						OpenTd "thead", "width=27% align=center"
							Response.Write "Fecha proceso"
						CloseTd
						OpenTd "thead", "width=27% align=center"
							Response.Write "Fecha efectiva"
						CloseTd
						OpenTd "thead", "width=27% align=center"
							Response.Write "Número de Confirmación"
						CloseTd
						OpenTd "thead", "width=22% align=center"
							Response.Write "Procesado por"
						CloseTd
						OpenTd "thead", "align=center width=24%"
							Response.Write "Detalles"
						CloseTd
					CloseTr
					'Start recordset data
					RecCount = 0
					Do While Not rsHistory.EOF And RecCount < rsHistory.PageSize
						If IsNull(rsHistory(1)) Then
							EffDate = "Indefinida"
						Else
							EffDate = FormatDateTime(rsHistory(1),2)
						End If
						OpenForm "Detail", "post", "sa_detail.asp", "onSubmit=formValidation(this)"
							PlaceInput "Contract", "hidden", Contract, ""
							PlaceInput "Product", "hidden", Product, ""
							PlaceInput "Plan", "hidden", Plan, ""
							PlaceInput "EffDate", "hidden", EffDate, ""
							PlaceInput "ConfNum", "hidden", rsHistory(2), ""
						If (rsHistory.AbsolutePosition Mod 2) = 0 Then
							OpenTr "class=teven"
						Else
							OpenTr "class=todd"
						End If
						For I = 0 To rsHistory.Fields.Count 'Cols
							Select Case I
								Case 0
									OpenTd "tbody", "align=center"
										If IsNull(rsHistory(I)) Then
											Response.Write "Indefinida"
										Else
											Response.Write FormatDateTime(rsHistory(I),2)
										End If
									CloseTd
								Case 1
									OpenTd "tbody", "align=center"
										Response.Write EffDate
									CloseTd
								Case 2, 3
									OpenTd "tbody", "align=center nowrap"
										If IsNull(rsHistory(I)) Then
											Response.Write "N/A"
										Else
											Response.Write rsHistory(I)
										End If
									CloseTd
							End Select
						Next
								OpenTd "tbody", "align=center"
									PlaceInput "submit", "submit", "detalles", "class=sbttn"
								CloseTd
							CloseForm
						CloseTr
						CloseForm
						RecCount = RecCount + 1
						rsHistory.MoveNext
					Loop
					rsHistory.MoveLast
					OpenTr ""
						OpenForm "prev", "post", "", ""
						OpenTd "thead", "align=center nowrap valign=middle"
							If  CInt(PgCount) > 1 Then
									PlaceInput "submit","submit", "< Anterior", "class=sbttn"
									PlaceInput "Page", "hidden", PgCount - 1, ""
									PlaceInput "Contract", "hidden", Contract, ""
									PlaceInput "Product", "hidden", Product, ""
									PlaceInput "Plan", "hidden", Plan, ""
							Else
								Response.Write "&nbsp;"
							End If
						CloseTd
						CloseForm
						OpenTd "thead", "colspan=3 align=center nowrap"
							Response.Write "Pagina: " & PgCount & " de " & Pages
							Response.Write " - No. Total de registros: " & rsHistory.RecordCount
						CloseTd
						OpenForm "next", "post", "", ""
						OpenTd "thead", "align=center nowrap valign=middle"
							If  CInt(PgCount) < CInt(Pages) Then
									PlaceInput "submit","submit", "Siguiente >", "class=sbttn"
									PlaceInput "Page", "hidden", PgCount + 1, ""
									PlaceInput "Contract", "hidden", Contract, ""
									PlaceInput "Product", "hidden", Product, ""
									PlaceInput "Plan", "hidden", Plan, ""
							Else
								Response.Write "&nbsp;"' & PgCount & " < " & Pages & " = " & (PgCount < Pages)
							End If
						CloseTd
					CloseForm
					CloseTr
				CloseTable
				rsHistory.Close
			End If
		CloseTd
	CloseTr
CloseTable
Response.Write "<p>&nbsp;</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf
CloseBody
CloseHTML
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>