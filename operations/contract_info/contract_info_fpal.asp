<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:			contract_info_fpal.asp 19800
'Path:				contract_info/
'Created By:		J Carreño 2003/09/11
'Parameters:		User must be logged on
'					Session("docNumber")'
' se elminina plataforma de consejo consultar version anterior 
'						I&T - WTG 20080313 Inclusion de consulta para clientes core
'Returns:			FPAL contract information
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

Authorize 6,25

Dim Sql 'SQL Sentences holder
Dim adoConn 'Database Connection
Dim rstDescription 'Results recordset
dim objRst	' Strategist PAS recordset
Dim rs, cn
Dim arrContract
Dim arrHIstory
Dim arrAsset, arrAccounts
Dim arrStanding, Status
Dim PrimaBruta, arrInsurance
Dim objSkMtrust, Accounts
Dim I,J 'Asset & Standing Alloc counters
Dim Contract, Product, Plan, ClientId, DocType, Name, Phone, OU, Pas, PasName, Frase
Dim ProvFund
dim Reference  '<I&T - DMPC>

Dim IsLegalRole ' I&T - WTG
IsLegalRole = False

'====================
' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
'====================
Dim IsCoreDescription, arrComplement
IsCoreDescription = "No se Conoce"
'====================
' <I&T - WTG: ISCORE>
'====================

Pas = Request.Form("Pas")
PasName = Request.Form("PasName")
Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
ClientId = Request.Form("ClientId")
Session("name")=ClientId
DocType = Request.Form("DocType")
Name = Request.Form("Name")
Phone = Request.Form("Phone")
OU = Request.Form("OU")
Set adoConn = GetConnpipelineDB

'==========================================================================
''<I&T - DMPC 2009/03/31 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia única a partir del Contrato y el producto
'==========================================================================
 Reference = GetReferenciaUnica( Product, Contract)
'==========================================================================

write_dataLog Response.Status,"contract_info_fpal.asp","process_selection Loaded by: " & Session("sp_miLogin") & " selected  for the Contract " & Contract,ClientId,"ContratoRetirosPAC_GetForContrato " & Contract & " - ContratoRetirosPAC_GetForContrato " & Contract,"N/A","null","Consulta","N/A"

Session("seleccionContrato")="false"

Set rstDescription = Server.CreateObject("ADODB.Recordset")
If ClientId = "" Then ClientId = 0

'====================
' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
'====================
write_sp_log adoConn, 500, "Iscore : " & CStr(ClientId) + ":" & DocType, Contract, Product, Plan, ClientId, 0, "", "contract_info_mfund.asp " & _
"Loaded by " & Session("sp_miLogin")

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

write_sp_log adoConn, 19800, "sppl_GetDetallesContrato", Contract, Product, Plan, ClientId, 0, "", "contract_info_fpal.asp " & _
"Loaded by " & Session("sp_miLogin")
If IsArray(arrContract) Then
	'Get Status Description
	Sql = "spsp_GetStatusDescription '" & Product & _
	"', '" & arrContract(14,0) & "'"
	rstDescription.Open Sql, adoConn
	If rstDescription.BOF And rstDescription.EOF Then
		Status = arrContract(14,0)
	Else
		Status = rstDescription.Fields("descripcion")
	End If
	rstDescription.Close
	write_sp_log adoConn, 19800, "spsp_GetStatusDescription", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
	" at contract_info_fpal.asp"
Else
	Status = "N/A"
End If

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
	write_sp_log adoConn, 19800, "spem_GetFPOBAnterior", Contract, Product, Plan, ClientId, 0, "", "contract_info_fpob.asp " & _
	"- " & Session("sp_miLogin")
Else
	ProvFund = "Ninguno"
End If


'Get Standing Allocation
Sql = "spsp_GetStandingAllocation " & Contract & ", '" & Product & "'"
rstDescription.Open Sql, adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrStanding = 0
Else
	arrStanding = rstDescription.GetRows()
End If
rstDescription.Close
write_sp_log adoConn, 19800, "spsp_GetStandingAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
" at contract_info_fpal.asp"

'Get employeers history
Sql = "sppl_EmpleadorxClienteProductoHist " & ClientId & ", '" & Product & "', '" & DocType & "'"
rstDescription.Open Sql,adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrHistory = 0
Else
	arrHistory = rstDescription.GetRows()
End If
rstDescription.Close

write_sp_log adoConn, 19800, "sppl_EmpleadorxClienteProductoHist", Contract, Product, Plan, ClientId, 0, "", "contract_info_fpal.asp " & _
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
			PlaceInput "Pas", "hidden", Request.Form("Pas"), ""
		CloseForm
%>
		<script language=javascript>
			document.menu_left.submit();
		</SCRIPT>
<%
		'Reload Left Menu -- END
	End If
If IsArray(arrContract) Then
OpenTable "90%", "t_table align=center"
	OpenTr ""
		OpenTd "", ""
			OpenTable "100%", ""
				OpenTr ""
					OpenTd "thead", ""
						If IsNull(arrContract(11,0)) Then
							Response.Write "N/A"
						Else
							Response.Write "Información de la afiliación a " & FormatDateTime(arrContract(11,0), 2) & vbCrLf
						End If
					CloseTD
				CloseTR
			CloseTable
			OpenTable "100%", ""
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


				OpenTr ""
					OpenTd "tbody", "width=30%"
						Response.Write "Producto"
					CloseTd					
					OpenTd "tbody", "width=5%"
						Response.Write " : "
					CloseTd
					OpenTd "tbody", "width=15%"
					If Left(Plan,1) = "U" Then
						Response.Write "UniFund"
					Else
						Sql = "spem_GetDescriptionProduct " & arrContract(0,0)
						rstDescription.Open Sql, adoConn
							Response.Write rstDescription.Fields("Descripcion")
						CloseTd
						rstDescription.Close
					End If
					OpenTd "tbody", "width=50%"
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						'Response.Write "Afiliación" '<I&T - DMPC - Modificado por Proceso de Referencia Unica - Recaudos>
						Response.Write "Contrato"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", ""
						'Response.Write arrContract(1,0)
						'<I&T - DMPC 2009/03/31 - Modificado por proceso de Referencia Única - Recaudos>
						if IsNull(Reference) or len(Reference)>12 Then
							Response.Write arrContract(1,0)
						else
							Response.Write Reference
						end if
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Estado"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", ""
						Response.Write Status
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Plan"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", ""
						Response.Write Plan
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Financial Planner"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", "nowrap"
						Response.Write  arrContract(36,0)
					CloseTd
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd										
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Fecha de Apertura Afiliación"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd
					OpenTd "tbody", ""
						If IsNull(arrContract(6,0)) Then
							Response.Write "N/A"
						Else
							Response.Write FormatDateTime(arrContract(6,0),2)
						End If
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd										
				CloseTr

				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Fecha Efectiva Afiliación"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd
					OpenTd "tbody", ""
						If IsNull(arrContract(7,0)) Then
							Response.Write "N/A"
						Else
							Response.Write FormatDateTime(arrContract(7,0),2)
						End If
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd										
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Proviene de"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd
					OpenTd "tbody", ""
							Response.Write ProvFund
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd										
				CloseTr


				OpenTr ""
					OpenTd "thead", ""
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead", ""
						Response.Write "Valores"
					CloseTd
				CloseTr
				
				'OpenTr ""
				'	OpenTd "tbody", ""
				'		Response.Write "Bono Pensional"
				'	CloseTd
				'	OpenTd "tbody", ""
				'		Response.Write ":"
				'	CloseTd
				'	OpenTd "tbody", ""
				'			Response.Write "Pendiente *" 
				'	CloseTd
				'	OpenTd "tbody", ""
				'		Response.Write "&nbsp;"
				'	CloseTd										
				'CloseTr	
							
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "Aportes"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd
					OpenTd "tbody", ""
						Dim Total
						Total = arrContract(9,0)
						if not isnull(Total) then
							Response.Write FormatCurrency(Total,2)
						else
							Response.Write FormatCurrency(0,2)
						end if
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd										
				CloseTr

				OpenTr ""
					OpenTd "thead", ""
						Response.Write "Total"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd
					OpenTd "tbody", ""
						if not isnull(Total) then
							Response.Write FormatCurrency(Total,2)
						else
							Response.Write FormatCurrency(0,2)
						end if
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd										
				CloseTr					

				OpenTr ""
					OpenTd "tbody", " colspan=4"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				'OpenTr ""
				'	OpenTd "tbody", " colspan=4"
				'		Response.Write "* Con Corte a: Pendiente"
				'	CloseTd
				'CloseTr
												
			'CloseTable
' get legalrole

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

			Response.Write "<br>"		
			'Build FPAL Information
			If RTrim(arrContract(0,0)) = "FPAL" Then
				OpenTable "100%", ""
					OpenTr ""
						OpenTd "thead", ""
						CloseTd
					CloseTr
					OpenTr ""
						OpenTd "thead", ""
							Response.Write "Composición actual de inversiones"
						CloseTd
					CloseTr
				CloseTable
				'Get Asset Allocation
				Sql = "spem_GetAssetAllocation '" & RTrim(arrContract(0,0)) & "', " & arrContract(1,0)
				rstDescription.Open Sql, adoConn
				If rstDescription.BOF And rstDescription.EOF Then
					arrAsset = 0
				Else
					arrAsset = rstDescription.GetRows()
				End If
				rstDescription.Close
				write_sp_log adoConn, 19800, "spem_GetAssetAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
				" at contract_info_fpal.asp"
				If IsArray(arrAsset) Then
				OpenTable "100%", "'' border=1"
					OpenTr "class=teven"
						OpenTd "thead", "width=25% align=center"
							Response.Write "Fondo"
						CloseTd
						OpenTd "thead", "width=25% align=center"
							Response.Write "Participación (%)"
						CloseTd
						OpenTd "thead", "width=25% align=center"
							Response.Write "No. de unidades"
						CloseTd
						OpenTd "thead", "width=25% align=center"
							Response.Write "Valor unidad ($)"
						CloseTd
						OpenTd "thead", "width=25% align=center"
							Response.Write "Saldo actual"
						CloseTd
					CloseTr
				For J = 0 To Ubound(arrAsset, 2) 'Rows
					If (J Mod 2) = 0 Then
						OpenTr " class=todd"
					Else
						OpenTr "class=teven"
					End If
					For I = 0 To UBound(arrAsset)-3
						If I <> 3 Then
							OpenTd "tbody", "align=center"
						Else
							OpenTd "tbody", "align=Right"
						End If
						Select Case I
							Case 1
								Response.Write FormatNumber(arrAsset(I,J),2) & "&nbsp;"
							Case 3' Get units value
								If IsNull(arrContract(11,0)) Then
									Sql = "sppl_getvalorhistund '" & FormatDateTime(Now(), 2) & "','" & _
									FormatDateTime(Now(), 2) & "','" & arrAsset(I+2,J) & "'"
								Else
									Sql = "sppl_getvalorhistund '" & FormatDateTime(arrContract(11,0), 2) & "','" & _
									FormatDateTime(arrContract(11,0), 2) & "','" & arrAsset(I+2,J) & "'"
								End If
								rstDescription.Open Sql,adoConn
								If rstDescription.BOF And rstDescription.EOF Then
									Response.Write "<p align=center>N/A</p>"
								Else
									Response.Write formatcurrency(rstDescription.Fields("valorunidad"), 6) & "&nbsp;"
								End If
								rstDescription.Close
								write_sp_log adoConn, 19800, "sppl_getvalorhistund", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
								" at contract_info_fpal.asp"
							Case 4
								If IsNull(arrAsset(I,J)) Then
									Response.Write "<p align=center>N/A</p>"
								Else
									Response.Write formatcurrency(arrAsset(I,J), 2) & "&nbsp;"
								End If
							Case Else
								Response.Write arrAsset(I,J) & "&nbsp;"
						End Select
						CloseTd
					Next
					CloseTr
				Next
				CloseTable
				Else
					OpenTable "90%","border=1"
						OpenTr "class=teven"
							OpenTd "thead","align=center"
								Response.Write "No hay composición actual de inversiones"
							CloseTd
						CloseTr
					CloseTable
				End If
				'Standing Allocation Table --- START
				OpenTable "50%", "'' border=0"
					OpenTr "class="
						OpenTd "'thead'", ""
							Response.Write "&nbsp;"
						CloseTd
					CloseTr
					OpenTr "class="
						OpenTd "'thead'", ""
							Response.Write "Standing Allocation"
						CloseTd
					CloseTr
				CloseTable
				OpenTable "50%", "'' border=1"
					If IsArray(arrStanding) Then
						OpenTr "class=teven"
							OpenTd "'thead'", " align=center"
								Response.Write "Fondo"
							CloseTd
							OpenTd "'thead'", " align=center"
								Response.Write "Porcentaje"
							CloseTd
						CloseTr
						For I = 0 To UBound(arrStanding, 2)
							If I Mod 2 = 0 Then
								OpenTr "class=todd"
							Else
								OpenTr "class=teven"
							End If
								OpenTd "''", " align=center"
									Response.Write arrStanding(0,I)
								CloseTd
								OpenTd "'money'", " align=center"
									Response.Write FormatPercent(CDbl(arrStanding(1,I))/100, 2)
								CloseTd
							CloseTr
						Next
					CloseTable
					OpenTable "", ""	
						OpenTr "class=tbody"
							OpenTd "''", "colspan=2"
								Response.Write "&nbsp;"
							CloseTd
						CloseTr
					Else
						OpenTr "class=tbody"
							OpenTd "''", "colspan=2"
								Response.Write "No hay Standing Allocation"
							CloseTd
						CloseTr
					End If
				CloseTable
				'Standing Allocation Table --- END
			End If
			
			'agregar información histórica de empleadores			
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
Else
	Response.Write "No hay detalles para este contrato"
End If
Response.Write "<P></P>"
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