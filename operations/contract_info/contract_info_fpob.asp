<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		contract_info_fpob.asp  10900
'Path:				contract_info/
'Created By:		Andres Felipe Orozco 2001/05/29
'Modified:			Fabio Calvache Julio 31 2003
'					Andres Felipe Orozco  2001/06/19
'					Guillermo Aristizabal  2001/07/28
'					Guillermo Aristizabal  2001/09/18 auth & log
'						A. Orozco 2001/10/08
'						Guillermo Aristizabal 2001/10/11
'						A. Orozco 2001/10/25
'						A. Orozco 2001/12/17
'						R. Lagos  2003/01/09 Add Metaname Contract
'						R. Lagos  2003/03/12 chage text Mora by Valor recibido por mora
'				Se borro plataforma de consejo definitivamente 
'						I&T - WTG 20080313 Inclusion de consulta para clientes core
'Parameters:		User must be logged on
'						Session("docNumber")
'Returns:			List of active contracts for the client
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

 
Authorize 5,14

Dim Sql 'SQL Sentences holder
Dim adoConn 'Database Connection
Dim rstDescription 'Results recordset
Dim rs, cn, ProdName, ProvFund
Dim arrContract, arrHIstory, arrSaldoFondos, arrDistribucionFutura
Dim arrAsset
Dim arrStanding
Dim I,J 'Asset & Standing Alloc counters
Dim Contract, Product, Plan, ClientId, DocType, Name, Phone, OU
Dim IsLegalRole ' I&T - WTG
IsLegalRole = False
dim Reference '<I&T - DMPC>

'====================
' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
'====================
Dim IsCoreDescription, arrComplement
IsCoreDescription = "No se Conoce"
'====================
' <I&T - WTG: ISCORE>
'====================

Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
Name = Request.Form("Name")
ClientId = Request.Form("ClientId")
Session("name")=ClientId
DocType = Request.Form("DocType")
Phone = Request.Form("Phone")
OU = Request.Form("OU")
Set adoConn = GetConnpipelineDB





Set rstDescription = Server.CreateObject("ADODB.Recordset")
'====================
' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
'====================
write_sp_log adoConn, 500, "Iscore : " & CStr(ClientId) + ":" & DocType, Contract, Product, Plan, ClientId, 0, "", "contract_info_mfund.asp " & _
"Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"contract_info_fpob.asp","process_selection Loaded by: " & Session("sp_miLogin") & " selected  for the Contract " & Contract,ClientId,"ContratoRetirosPAC_GetForContrato " & Contract & " - ContratoRetirosPAC_GetForContrato " & Contract,"N/A","null","Consulta","N/A"

Session("seleccionContrato")="false"

Set rstDescription = Server.CreateObject("ADODB.Recordset")

Sql = "spsp_ComplemnetData_GetByClient '" & CStr(DocType) & "', " & CStr(ClientId)
rstDescription.Open Sql, adoConn

If rstDescription.BOF And rstDescription.EOF Then
	arrComplement = 0
	IsCoreDescription = "No se Conoce"
Else
	arrComplement = rstDescription.GetRows()

	If IsArray(arrComplement) Then
		IsCoreDescription = arrComplement(0,0)	
	End If
End If
rstDescription.Close
'====================
' <I&T - WTG: ISCORE>
'====================

'==========================================================================
''<I&T - DMPC 2009/02/20 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia única a partir del Contrato y el producto
'==========================================================================
 Reference = GetReferenciaUnica( Product, Contract)
'==========================================================================

'Get LegalRole
Sql = "sppl_GetLegalRole '" & DocType & "', " & ClientId & ", " & RTrim(Contract) & ", '" & RTrim(Product) & "', '" & RTrim(Plan) &"'"
rstDescription.Open Sql, adoConn

If rstDescription.BOF And rstDescription.EOF Then
	IsLegalRole = False
Else	
	If CInt(rstDescription.Fields("COUNTLEGALROLE")) <> 0 Then
		IsLegalRole = True
	Else
		IsLegalRole = False
	End If
End If
rstDescription.Close


'Get Contract description
Sql = "sppl_GetDetallesContrato " & Contract & _
", '" & Product & "', '" & Plan & "'"
'Response.Write Sql 
'Response.End 
rstDescription.Open Sql, adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrContract = 0
Else
	arrContract = rstDescription.GetRows()
End If
rstDescription.Close



if  IsArray(arrContract) then
	if not (isnull(arrContract(15,0)) or (isnull(arrContract(35,0)))) then
			AuthorizeContractAccess Cstr(Session("sp_AccessLevel")),  CStr(Session("sp_IdAgte")), _
						CStr(arrContract(15,0)), CStr(arrContract(35,0)) 
	end if
end if



write_sp_log adoConn, 10900, "sppl_GetDetallesContrato", Contract, Product, Plan, ClientId, 0, "", "contract_info_fpob.asp " & _
"Loaded by " & Session("sp_miLogin")

If not IsArray(arrContract) Then
	Response.Write "No hay detalles para este contrato"
	Response.End 
end if


'Get Product description
Sql = "spem_GetDescriptionProduct " & arrContract(0,0)
rstDescription.Open Sql, adoConn
ProdName = rstDescription.Fields("Descripcion")
rstDescription.Close



write_sp_log adoConn, 10900, "spem_GetDescriptionProduct", Contract, Product, Plan, ClientId, 0, "", "contract_info_fpob.asp " & _
"- " & Session("sp_miLogin")
'Get Previous Fund
If arrContract(32,0) <> "" and arrContract(33,0) <> "" Then
	Sql = "spem_GetFPOBAnterior '" & arrContract(32,0) & "', " & arrContract(33,0)
	rstDescription.Open Sql, adoConn
	If rstDescription.BOF And rstDescription.EOF Then
		ProvFund = "N/A"
	Else
		ProvFund = rstDescription.Fields("razonsocial")
	End If
	rstDescription.Close
	write_sp_log adoConn, 10900, "spem_GetFPOBAnterior", Contract, Product, Plan, ClientId, 0, "", "contract_info_fpob.asp " & _
	"- " & Session("sp_miLogin")
Else
	ProvFund = "Ninguno"
End If
'Get funds history
Sql = "sppl_EmpleadorxClienteProductoHist " & ClientId & ", '" & Product & "', '" & DocType & "'"
rstDescription.Open Sql,adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrHistory = 0
Else
	arrHistory = rstDescription.GetRows()
End If
rstDescription.Close

'Get Composición actual 
Sql = "spem_GetAssetAllocation " & Product & " ," & Contract
rstDescription.Open Sql,adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrSaldoFondos = 0
Else
	arrSaldoFondos = rstDescription.GetRows()
End If
rstDescription.Close


'Get Distribución Furtura
Sql = "spsp_GetStandingAllocation " & Contract & " ," & Product & " ,'" & Plan & "'"
rstDescription.Open Sql,adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrDistribucionFutura = 0
Else
	arrDistribucionFutura = rstDescription.GetRows()
End If
rstDescription.Close


write_sp_log adoConn, 10900, "sppl_EmpleadorxClienteProductoHist", Contract, Product, Plan, ClientId, 0, "", "contract_info_fpob.asp " & _
"- " & Session("sp_miLogin")

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	OpenBody "", ""
	If InStr(1, Request.ServerVariables("HTTP_REFERER"), "menu.asp") = 0 Then
		'Reload Left Menu -- START
		OpenForm "menu_left", "post", "../menu/menu.asp", "target=menu"
			PlaceInput "Name", "hidden", Request.Form("Name"), ""
			PlaceInput "ClientId", "hidden", ClientId, ""
			PlaceInput "DocType",  "hidden", DocType, ""
			PlaceInput "Phone", "hidden", Phone, ""
			PlaceInput "Contract", "hidden", Contract, ""
			PlaceInput "Product", "hidden", Product, ""
			PlaceInput "Plan", "hidden", Plan, ""
			PlaceInput "OU", "hidden", OU, ""
		CloseForm
%>
		<script language=javascript>
			document.menu_left.submit();
		</SCRIPT>
<%
		'Reload Left Menu -- END
	End If
OpenTable "90%","'' align=center"
	OpenTr ""
		OpenTd "","tbody"
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Información de la afiliación a " & _
						FormatDateTime(arrContract(11,0),2) & "<br>"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%",""
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Cliente"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", "nowrap"
						Response.Write Name
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Identificación"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", "nowrap"
						Response.Write ClientId
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				
'==========================================================================
' Add Document type abril 10 2003  Fabio Calvache
'==========================================================================

				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Tipo de Identificación"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", "nowrap"
						Response.Write arrcontract(4, 0)
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr

'====================
' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
'====================
				OpenTr ""
					OpenTd "teven", " style=""FONT-WEIGHT: bold"""
						Response.Write "Segmento"
					CloseTd
					OpenTd "teven", " style=""FONT-WEIGHT: bold"""
						Response.Write ":"
					CloseTd					
					OpenTd "teven", "nowrap style=""FONT-WEIGHT: bold"""
						Response.Write IsCoreDescription
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
'====================
' <I&T - WTG: ISCORE>
'====================

'==========================================================================
' End add Document type abril 10 2003  Fabio Calvache
'==========================================================================
				
				OpenTr ""
					OpenTd "tbody","width=40%"
						Response.Write "Producto"
					CloseTd
					OpenTd "tbody","width=5%"
						Response.Write ":"
					CloseTd					
					OpenTd "tbody","width=25%"
						Response.Write ProdName
					CloseTd
					OpenTd "thead","width=30%"
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						'Response.Write "Afiliación" '<I&T - DMPC 2009/02/20 - Modificado por proceso de Referencia Unica - Recaudos>
						Response.Write "Contrato"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd										
					OpenTd "tbody",""
						'Response.Write Contract '<I&T - DMPC 2009/02/20 - Modificado por proceso de Referencia Unica - Recaudos>
						response.Write Reference
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd				
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Plan"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd										
					OpenTd "tbody",""
						Response.Write Plan
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd				
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Financial Planner"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd					
					
					OpenTd "tbody","nowrap"
						Response.Write arrContract(36,0)
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd				
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Fecha de apertura afiliación"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd										
					OpenTd "tbody",""
						Response.Write FormatDateTime(arrContract(6,0),2)
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd				
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Fecha efectiva de afiliación"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd					
					
					OpenTd "tbody",""
						Response.Write FormatDateTime(arrContract(7,0),2)
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd				
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Proviene de"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd										
					OpenTd "tbody",""
						Response.Write ProvFund
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd				
				CloseTr
' Para informar el objetivo del contrato buscndo en el metaname del contrato				
				if  (ltrim(arrContract(23,0))) = "" or IsNull(arrContract(23,0))then
				else
					OpenTr ""
						OpenTd "tbody", ""
							Response.Write "Objetivo Contrato "
						CloseTd
						OpenTd "tbody", ""
							Response.Write ":"
						CloseTd
						OpenTd "tbody", ""
							Response.Write arrContract(23,0)
						CloseTd
						OpenTd "tbody", ""
							Response.Write "&nbsp;"
						CloseTd										
					CloseTr
				END IF
' Fin del objetivo del contrato
				
			CloseTable
			OpenTable "100%",""
				OpenTr ""
					OpenTd "",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Valores"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%",""
				OpenTr ""
					OpenTd "tbody","width=40%"
						Response.Write "Saldo obligatorio"
					CloseTd
					OpenTd "tbody","width=5%"
						Response.Write ":"
					CloseTd										
					OpenTd "tbody","width=20% align=right"
						Response.Write FormatCurrency(arrContract(28,0),2)
					CloseTd
					OpenTd "tbody","width=35%% align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Valor recibido por mora"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd															
					OpenTd "tbody","align=right"
						Response.Write FormatCurrency(arrContract(31,0),2)
					CloseTd
					OpenTd "tbody","align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Saldo voluntario empleador"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd										
					OpenTd "tbody","align=right"
						Response.Write FormatCurrency(arrContract(30,0),2)
					CloseTd
					OpenTd "tbody","align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Saldo voluntario afiliado"
					CloseTd
					OpenTd "tbody",""
						Response.Write ":"
					CloseTd										
					OpenTd "tbody","align=right"
						Response.Write FormatCurrency(arrContract(29,0),2)
					CloseTd
					OpenTd "tbody","align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "balance",""
						Response.Write "Total"
					CloseTd
					OpenTd "balance",""
						Response.Write ":"
					CloseTd															
					OpenTd "balance","align=right"
						Dim Total
						Total = arrContract(28,0)+arrContract(29,0)+arrContract(30,0)+arrContract(31,0)
						Response.Write FormatCurrency(Total,2)
					CloseTd
					OpenTd "tbody","align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			CloseTable
			'getlegalrole
			If IsLegalRole Then
				OpenTable "100%", "aling='center'"
					OpenTr ""
						OpenTd "", ""
							Dim strOpen 
							Dim sendData
							sendData = "6;2|7;" & DocType & "|8;" & ClientId & "|11;" & Contract & "|12;" & RTrim( Product)
							sendData = "@AppliedFilter=" & sendData
							sendData = Application("LegalRoleReport") & sendData						
							strOpen = "JavaScript:location.href ='" & sendData & "';"
							PlaceInput "btnLegalRole", "Button", "Ver Roles Legales", " class='sbttn' OnClick=""" & strOpen & """"
						CloseTd
					CloseTr
				CloseTable
			End If
			
			'Composición actual de inversiones
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "tbody","align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Composición Actual de inversiones"	
					CloseTd
				CloseTr
			CloseTable
			If IsArray(arrSaldoFondos) Then
				OpenTable "100%","'' align=center border=1"
					OpenTr "class=teven2 align=center"
						OpenTd "thead",""
							Response.Write "Fondo"
						CloseTd
						OpenTd "thead",""
							Response.Write "Participación (%)"
						CloseTd
						OpenTd "thead",""
							Response.Write "No de unidades"
						CloseTd
						OpenTd "thead",""
							Response.Write "Valor unidad ($)"
						CloseTd
						OpenTd "thead",""
							Response.Write "Saldo Actual"
						CloseTd
					CloseTr
					For J = 0 To UBound(arrSaldoFondos,2)
						If (J Mod 2) = 0 Then
							OpenTr "class=todd align=center"
						Else
							OpenTr "class=teven align=center"
						End If
						For I = 0 To UBound(arrSaldoFondos) - 3
							OpenTd "tbody",""
							If Not(IsNull(arrSaldoFondos(I,J))) THEN
								Select Case I
									Case 0
										Response.Write arrSaldoFondos(I,J)
									Case 1
										Response.Write arrSaldoFondos(I,J) & "%"
									Case 2
										Response.Write arrSaldoFondos(I,J)
									Case 3,4
										Response.Write FormatCurrency(arrSaldoFondos(I,J),2)
								    
								End Select
							End If
							CloseTd
						Next
						CloseTr
					Next
				CloseTable
			else
				OpenTable "100%","'' align=center"
					OpenTr "class=teven2 align=center"
						OpenTd "thead","width=100%"
							Response.Write "No hay composición actual de inversiones"
						CloseTd
					CloseTr
				CloseTable
			End If
			
			'Distribución de futuros aportes
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "tbody","align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Distribución de futuros aportes"	
					CloseTd
				CloseTr
			CloseTable
			If IsArray(arrDistribucionFutura) Then
				OpenTable "100%","'' align=center border=1"
					OpenTr "class=teven2 align=center"
						OpenTd "thead",""
							Response.Write "Fondo"
						CloseTd
						OpenTd "thead",""
							Response.Write "Porcentaje (%)"
						CloseTd
						OpenTd "thead",""
							Response.Write "Fecha Selección del afiliado"
						CloseTd
					CloseTr
					For J = 0 To UBound(arrDistribucionFutura,2)
						If (J Mod 2) = 0 Then
							OpenTr "class=todd align=center"
						Else
							OpenTr "class=teven align=center"
						End If
						For I = 0 To UBound(arrDistribucionFutura)
							If I <> 2 Then
								OpenTd "tbody",""
								If Not(IsNull(arrDistribucionFutura(I,J))) THEN
									Select Case I
										Case 0
											Response.Write arrDistribucionFutura(I,J)
										Case 1
											Response.Write arrDistribucionFutura(I,J) & "%"
										Case 3
											If Not(IsNull(arrDistribucionFutura(I,J))) Then
												Response.Write arrDistribucionFutura(I,J)
											Else
												Response.Write "&nbsp;"
											End If
									End Select
								End If
								CloseTd
							End If
						Next
						CloseTr
					Next
				CloseTable
			else
				OpenTable "100%","'' align=center"
					OpenTr "class=teven2 align=center"
						OpenTd "thead","width=100%"
							Response.Write "No hay distribución de futuros aportes"
						CloseTd
					CloseTr
				CloseTable
			End If
			
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "tbody","align=right"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Información histórica de empleadores"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%","'' align=center border=1"
				OpenTr "class=teven align=center"
					OpenTd "thead","width=25%"
						Response.Write "Nombre empleador"
					CloseTd
					OpenTd "thead","width=25%"
						Response.Write "Documento"
					CloseTd
					OpenTd "thead","width=25%"
						Response.Write "Fecha inicial"
					CloseTd
					OpenTd "thead","width=25%"
						Response.Write "Fecha final"
					CloseTd
				CloseTr
				If IsArray(arrHIstory) Then
				For J = 0 To UBound(arrHistory,2)
					If (J Mod 2) = 0 Then
						OpenTr "class=todd align=center"
					Else
						OpenTr "class=teven align=center"
					End If
					For I = 0 To UBound(arrHistory)
						OpenTd "tbody",""
						Select Case I
							Case 0
								Response.Write arrHistory(1,J)
							Case 1
								Response.Write arrHistory(0,J)
							Case 2,3
								If Not(IsNull(arrHistory(I,J))) Then
									Response.Write FormatDateTime(arrHistory(I,J),2)
								Else
									Response.Write "&nbsp;"
								End If
						End Select
						CloseTd
					Next
					CloseTr
				Next
				End If
			CloseTable
		CloseTd
	CloseTr
CloseTable
CloseBody
CloseHTML
Set rstDescription = Nothing

CloseConnpipelineDB
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
