<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		contract_info_others.asp 12100
'Path:				contract_info/
'Created By:		A. Orozco 2001/07/26
'Last Modified:		Fabio Calvache Abril 10 2003
' Add Document type abril 10 2003  Fabio Calvache
'					A. Orozco 2001/09/10
'					Guillermo Aristizabal  2001/09/18 auth & log
'						A. Orozco 2001/10/08
'						Guillermo Aristizabal 2001/10/11
'						A. Orozco 2001/10/25
'						APC	2001/11/07 description validation prodname
'						Se elimina definitivamente los cambios de plataforma de consejo cualquier cambio verificar version anterior
'						I&T - WTG 20080313 Inclusion de consulta para clientes core
'						I&T - WTG 20090210 Inclusión de cambio nivel de servicio
'Parameters:		User must be logged on
'						Session("docNumber")
'Modificado Por Julian Zapata 2017-02-02
			'Validar Autorizacion cuando la Sociedad del Worker que ingresa es del Canal Intermediario

'Parameters:		User must be logged on
'						Session("docNumber")
'			Julian Zapata Desarrollo-IT 02-02-2017 Se modifica para que una Sociedad AliadoEstrategico y/o Promotora pueda ver el detalle de los contratos de las sociedades asociadas
'				 Esto para la estructura de Sociedades Canal Intermediario.
'Returns:			contract information for products different from MFUND or FPOB
'Additional Information:
'===================================================================================

'Option Explicit
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
Authorize 1,16

Dim Sql 'SQL Sentences holder
Dim adoConn 'Database Connection
Dim rstDescription 'Results recordset
Dim rs, cn, ProdName, ProvFund
Dim arrContract, arrHIstory
Dim arrAsset
Dim arrStanding
Dim Contract, Product, Plan, ClientId, DocType, Name, Phone
Dim I,J 'Asset & Standing Alloc counters

dim Reference '<I&T - DMPC>

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

Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
Name = Request.Form("Name")
ClientId = Request.Form("ClientId")
Session("name")=ClientId
DocType = Request.Form("DocType")
Phone = Request.Form("Phone")
'Connect to database
Set adoConn = GetConnpipelineDB
Set rstDescription = Server.CreateObject("ADODB.Recordset")

'====================
' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
'====================
write_sp_log adoConn, 500, "Iscore : " & CStr(ClientId) + ":" & DocType, Contract, Product, Plan, ClientId, 0, "", "contract_info_mfund.asp " & _
"Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"contract_info_other.asp","process_selection Loaded by: " & Session("sp_miLogin") & " selected  for the Contract " & Contract,ClientId,"ContratoRetirosPAC_GetForContrato " & Contract & " - ContratoRetirosPAC_GetForContrato " & Contract,"N/A","null","Consulta","N/A"


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
''<I&T - DMPC 2009/02/13 - Modificado por proceso de Referencia Unica - Recaudos>
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
rstDescription.Open Sql, adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrContract = 0
Else
	arrContract = rstDescription.GetRows()
End If
rstDescription.Close

'=============================================================================================
			'Modificado Por Julian Zapata 2017-02-02
			'Validar Autorizacion cuando la Sociedad del Worker que ingresa No es del Canal intermediario
			'=============================================================================================
if  IsArray(arrContract) AND Session("esAgenteIntermediario") = 0 then
	if not (isnull(arrContract(15,0)) or (isnull(arrContract(35,0)))) then
			AuthorizeContractAccess Cstr(Session("sp_AccessLevel")),  CStr(Session("sp_IdAgte")), _
						CStr(arrContract(15,0)), CStr(arrContract(35,0)) 
	end if
end if

	'=============================================================================================
			'Modificado Por Julian Zapata 2017-02-02
			'Validar Autorizacion cuando la Sociedad del Worker que ingresa es del Canal Intermediario
			'=============================================================================================
if  IsArray(arrContract) AND Session("esAgenteIntermediario") = 1 then
	if not (isnull(arrContract(35,0))) then
		 if CStr(Session("sp_idSoc")) = CStr(arrContract(109,0)) then			
		   IdSociedadIntermediaria = CStr(arrContract(109,0))		   
		 else 
			if CStr(Session("sp_idSoc")) = CStr(arrContract(108,0)) then
				IdSociedadIntermediaria = CStr(arrContract(108,0))
			else 
				IdSociedadIntermediaria = CStr(arrContract(35,0))
			end if
		 end if
    AuthorizeContractAccess Cstr(Session("sp_AccessLevel")),  CStr(Session("sp_IdAgte")), _
						CStr(arrContract(15,0)), IdSociedadIntermediaria 
	end if
end if


'write_sp_log(connection, page_id, sp, contract, product, plan, client_id, error, conf_num, text)
write_sp_log adoConn, 12100, "sppl_GetDetallesContrato", Contract, Product, Plan, ClientId, 0, "", "contract_info_other.asp " & _
"- Loaded by" & Session("sp_miLogin")

If IsArray(arrContract) Then
	'Get Product description
	Sql = "spem_GetDescriptionProduct " & arrContract(0,0)
	rstDescription.Open Sql, adoConn
	If rstDescription.BOF And rstDescription.EOF Then
		ProdName = arrContract(0,0)
	Else
		ProdName = rstDescription.Fields("Descripcion")
	End if
	rstDescription.Close

	'write_sp_log(connection, page_id, sp, contract, product, plan, client_id, error, conf_num, text)
	write_sp_log adoConn, 12100, "spem_GetDescriptionProduct", Contract, Product, Plan, ClientId, 0, "", "contract_info_other.asp " & _
	"- Loaded by" & Session("sp_miLogin")
End If

'Consulta de la información de Assets -  Inicio
If RTrim(arrContract(0,0)) = "FCES" Then
	'Get Asset Allocation
	Sql = "spem_GetAssetAllocation '" & RTrim(arrContract(0,0)) & "', " & arrContract(1,0)
	rstDescription.Open Sql, adoConn
	If rstDescription.BOF And rstDescription.EOF Then
		arrAsset = 0
	Else
		arrAsset = rstDescription.GetRows()
	End If
	rstDescription.Close
	write_sp_log adoConn, 500, "spem_GetAssetAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_other.asp"
End If
'Consulta de la información de Assets -  Fin
				
	Set rstDescription = Nothing
'Close Database connetction
CloseConnpipelineDB

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	OpenBody "''", "bgcolor='#FFFFFF' text='#000000'"
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
'			PlaceInput "OU", "hidden", OU, ""
		CloseForm
		
%>
		<script language=javascript>
			document.menu_left.submit();
		</SCRIPT>
<%
		'Reload Left Menu -- END
	End If
If IsArray(arrContract) Then
OpenTable "90%","'' align=center"
	OpenTr ""
		OpenTd "",""
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
					OpenTd "tbody", "width=5%"
						Response.Write ":"
					CloseTd										
					OpenTd "tbody","width=15%"
						Response.Write ProdName
					CloseTd
					OpenTd "tbody","width=50%"
						Response.Write "&nbsp;"
					CloseTd
					
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						'Response.Write "Afiliación" '<I&T - DMPC 2009/02/09 - Modificado por proceso de Referencia Unica - Recaudos>
						Response.Write "Contrato"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd															
					OpenTd "tbody",""
					'	Response.Write Contract '<I&T - DMPC 2009/02/09 - Modificado por proceso de Referencia Unica - Recaudos>
					if IsNull(Reference) Then
						response.Write Contract
					else
						response.Write Reference
						end if
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "tbody",""
						Response.Write "Financial Planner"
					CloseTd
					OpenTd "tbody", ""
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
						Response.Write "Fecha de ingreso"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd										
					OpenTd "tbody",""
						Response.Write FormatDateTime(arrContract(6,0),2)
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				
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
				
				if  (ltrim(arrContract(23,0))) = "" or IsNull(arrContract(23,0)) then
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
				
						
				
				OpenTr ""
					OpenTd "tbody", "colspan=4"
						Response.Write "&nbsp;"
					CloseTd															
				CloseTr

				OpenTr ""
					OpenTd "balance",""
						Response.Write "Saldo"
					CloseTd
					OpenTd "balance", ""
						Response.Write ":"
					CloseTd										
					OpenTd "balance",""
						Response.Write FormatCurrency(arrContract(9,0),2)
					CloseTd
					OpenTd "balance",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			CloseTable
		CloseTd
	CloseTr
	
CloseTable

	
Else
	OpenTable "90%", "'' border=1"
		OpenTr "class=teven"
			OpenTd "thead", "align=center"
				Response.Write "No Hay Detalles"
			CloseTd
		CloseTr
	CloseTable
End If

'	'I&T - DMPC - 25/11/2009 - Inclusión de Asset y Standing por Reforma Financiera
			If RTrim(arrContract(0,0)) = "FCES" Then
				OpenTable "100%","'' align=center"
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
'				'Get Asset Allocation
'				Sql = "spem_GetAssetAllocation '" & RTrim(arrContract(0,0)) & "', " & arrContract(1,0)
'				rstDescription.Open Sql, adoConn
''				If rstDescription.BOF And rstDescription.EOF Then
'					arrAsset = 0
'				Else
'					arrAsset = rstDescription.GetRows()
'				End If
'				rstDescription.Close
'				write_sp_log adoConn, 500, "spem_GetAssetAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
'				" at contract_info_other.asp"
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
						OpenTd "thead", "width=25% align=center"
							Response.Write "Disp."
						CloseTd
						
					CloseTr
				For J = 0 To Ubound(arrAsset, 2) 'Rows
					If (J Mod 2) = 0 Then
						OpenTr " class=todd"
					Else
						OpenTr "class=teven"
					End If
					For I = 0 To UBound(arrAsset)
					 IF i <> 5 and I <> 6 Then
						If I <> 3 Then
							OpenTd "tbody", "align=center"
						Else
							OpenTd "tbody", "align=Right"
						End If
					  End If
						Select Case I
							Case 1
								Response.Write FormatNumber(arrAsset(I,J),2) '& "&nbsp;"
							Case 3' Get units value
								Set rstDescription = Server.CreateObject("ADODB.Recordset")
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
									Response.Write formatcurrency(rstDescription.Fields("valorunidad"), 6) '& "&nbsp;"
								End If
								rstDescription.Close
								write_sp_log adoConn, 500, "sppl_getvalorhistund", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
								" at contract_info_other.asp"
							Case 7
								If arrAsset(I,J)= false Then
									Response.Write "<p align=center><font color=red><a class=info href='#'>NO<span><strong>PRECAUCIÓN</strong>:Este fondo está cerrado</span></a></font></p>"
								Else
									Response.Write "SI"
								End If
							Case 5
									Response.Write ""
							Case 6

									Response.Write ""
							Case Else
								If arrAsset(7,J)= false Then
									Response.Write "<p align=center><font color=red><a class=info href='#'>" & arrAsset(I,J) & _
									"<span><strong>PRECAUCIÓN</strong>:Este fondo está cerrado</span></a></font></p>"
								Else
									'===============
									' I&T-WTG 20080522 Formato $ para el saldo disponible
									'===============
									If I = 4 Then
										Response.Write FormatCurrency(arrAsset(I,J))
									Else
										Response.Write arrAsset(I,J)
									End If
									'===============
									' I&T-WTG 20080522 
									'===============
								End If
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
'				'Standing Allocation Table --- START
'				OpenTable "50%", "'' border=0"
'					OpenTr "class="
'						OpenTd "'thead'", ""
'
'							Response.Write "&nbsp;"
'						CloseTd
'					CloseTr
'					OpenTr "class="
'						OpenTd "'thead'", ""
'							Response.Write "Standing Allocation"
'						CloseTd
'					CloseTr
'				CloseTable
'				OpenTable "50%", "'' border=1"
'					If IsArray(arrStanding) Then
'						OpenTr "class=teven"
'							OpenTd "'thead'", " align=center"
'								Response.Write "Fondo"
'							CloseTd
'							OpenTd "'thead'", " align=center"
'								Response.Write "Porcentaje"
'							CloseTd
'						CloseTr
'						For I = 0 To UBound(arrStanding, 2)
'							If I Mod 2 = 0 Then
'								OpenTr "class=todd"
'							Else
'								OpenTr "class=teven"
'							End If
'								OpenTd "''", " align=center"
'									Response.Write arrStanding(0,I)
'								CloseTd
'								OpenTd "'money'", " align=center"
'									Response.Write FormatPercent(CDbl(arrStanding(1,I))/100, 2)
'								CloseTd
'							CloseTr
'						Next
'					CloseTable
'					OpenTable "", ""	
'						OpenTr "class=tbody"
'							OpenTd "''", "colspan=2"
'								Response.Write "&nbsp;"
'							CloseTd
'						CloseTr
'					Else
'						OpenTr "class=tbody"
'							OpenTd "''", "colspan=2"
'								Response.Write "No hay Standing Allocation"
'							CloseTd
'						CloseTr
'					End If
'				CloseTable
'				'Standing Allocation Table --- END
end if
'I&T - DMPC - 25/11/2009 - Inclusión de Asset y Standing por Reforma Financiera - fin

CloseBody
CloseHTML
'Reload Left Menu -- START
OpenForm "menu", "post", "../menu/menu.asp", "target=menu"
	PlaceInput "Name", "hidden", Name, ""
	PlaceInput "ClientId", "hidden", ClientId, ""
	PlaceInput "DocType",  "hidden", DocType, ""
	PlaceInput "Contract", "hidden", Contract, ""
	PlaceInput "Product", "hidden", Product, ""
	PlaceInput "Phone", "hidden", Phone, ""
	PlaceInput "Plan", "hidden", Plan, ""
	PlaceInput "Option", "hidden", 0, ""
CloseForm
%>
<script language=javascript>
	document.menu.submit();
</SCRIPT>
<%
'Reload Left Menu -- END
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
