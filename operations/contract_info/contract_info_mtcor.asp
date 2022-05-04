<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		contract_info_mtcor.asp 500
'Path:				contract_info/
'Created By:		Andres Felipe Orozco 2001/11/07
'Modified:			Andres Felipe Orozco  2001/11/08
'Modified:			Andres Felipe Orozco  2001/11/27
'Modified:			Andres Felipe Orozco  2002/01/04
'Modified:			Andres Felipe Orozco  2002/01/28
'Modified:			Andres Felipe Orozco  2002/03/22
'Modified:			Fabio Calvache		  2002/10/23 - account_balance from as400
'Modified:			Rafael Lagos   2002/11/26 adicionar total impuestos, total retiros  ordenar la presentacion del saldo
'Modified:			Fabio Calvache	Add Document type abril 10 2003  Fabio Calvache
'Modified:			IT  se elimina la plataforma de consejo cualquier consulta versión anterior
'Modified:			I&T - WTG 20080313 Inclusion de consulta para clientes core
'Modified:			I&T - WTG 20090210 Inclusión de cambio nivel de servicio
'Parameters:		User must be logged on
'						Session("docNumber")
'Returns:			MTCOR contract information
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
Authorize 5,1

Dim Sql 'SQL Sentences holder
Dim adoConn 'Database Connection
Dim rstDescription, RstImpuest'Results recordset
Dim rs, cn
Dim arrContract
Dim arrAsset, arrAccounts
Dim arrStanding, Status, frase
Dim PrimaBruta
Dim arrSaldos400, objSkMtrust, InputDate
Dim GrowthParams, GrowthParamsInfo, Accounts, Beneficiaries, Terceros
Dim Thirds
Dim I,J 'Asset & Standing Alloc counters
Dim Contract, Product, Plan, ClientId, DocType, Name, Phone, OU, Pas
dim Pagina
dim ObjCon
dim Param, RetRealizados, TaxCobrados

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
ClientId = Request.Form("ClientId")
DocType = Request.Form("DocType")
Name = Request.Form("Name")
Session("name")=ClientId

write_dataLog Response.Status,"contract_info_mtcor.asp","process_selection Loaded by: " & Session("sp_miLogin") & " selected  for the Contract " & Contract,ClientId,"ContratoRetirosPAC_GetForContrato " & Contract & " - ContratoRetirosPAC_GetForContrato " & Contract,"N/A","null","Consulta","N/A"


Session("seleccionContrato")="false"
%>

Phone = Request.Form("Phone")
OU = Request.Form("OU")
Pas = Request.Form("Pas")

if day(date()) = 1 then
	InputDate = year( dateadd( "d" ,-1, date())  ) & Right(string(2, "00") & month( dateadd( "d",-1, date()) ), 2) & _
	Right(string(2, "0") & day( dateadd( "d",-1, date()) ), 2)
else
	InputDate = year(date()) & Right(string(2, "00") & month(date()), 2) & _
	Right(string(2, "0") & day(date()), 2)
end if

'Set adoConn = GetConnpipelineDB
'Set RstImpuest = Server.CreateObject("ADODB.Recordset")
'Sql = "transaction_Get_Debits_Tax_Mtcor " & Contract 
'RstImpuest.Open Sql, adoConn

'If RstImpuest.BOF And RstImpuest.EOF Then
'	RetRealizados =  0
'	TaxCobrados = 0
'	
'Else'
'	RetRealizados = RstImpuest.Fields(0)
'	TaxCobrados = RstImpuest.Fields(1)
'End If
'RstImpuest.Close



Set adoConn = GetConnpipelineDB

'=======================================================================================
'Start of AS400
'=======================================================================================
	
If Contract = "" Then
	Response.Write("Debe existir Número de Contrato.")
'	Param = Application("AccountBal") & "?Ctr=" & Contract & "&FP=P"
'	Pagina = ObjCon.urlReader(Param)

Else

	''SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNIONÓ EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN
	''Set ObjCon = Server.CreateObject("Conexion.conectar") 

	Param = Application("AccountBal") & "?Ctr=" & Contract & "&FP=P"

	Set ObjCon = CreateObject("MSXML2.XMLHTTP")
	ObjCon.Open "GET", Param , false
	ObjCon.setRequestHeader "Content-Type","text/xml"
	ObjCon.Send(null)
	pagina = ObjCon.responseText

	''Pagina = ObjCon.urlReader(Param)
	''FIN SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNIONÓ EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN

	
End If
set ObjCon = nothing
'=======================================================================================
'End of AS400
'=======================================================================================

'==========================================================================
''<I&T - DMPC 2009/02/13 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia única a partir del Contrato y el producto
'==========================================================================
 Reference = GetReferenciaUnica( Product, Contract)
'==========================================================================


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



'Get Contract description

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


write_sp_log adoConn, 500, "sppl_GetDetallesContrato", Contract, Product, Plan, ClientId, 0, "", "contract_info_mfund.asp " & _
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
	write_sp_log adoConn, 500, "spsp_GetStatusDescription", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
	" at contract_info_mfund.asp"
Else
	Status = "N/A"
End If

'Get Registered accounts
Sql = "sppl_CuentasRegXContrato " & Contract
rstDescription.Open Sql, adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrAccounts = 0
Else
	arrAccounts = rstDescription.GetRows()
End If
rstDescription.Close

write_sp_log adoConn, 500, "sppl_CuentasRegXContrato", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
" at contract_info_mfund.asp"

'Get Standing Allocation
Sql = "spsp_GetStandingAllocation " & Contract & ", '" & Product & "'"
rstDescription.Open Sql, adoConn
If rstDescription.BOF And rstDescription.EOF Then 
	arrStanding = 0
Else
	arrStanding = rstDescription.GetRows()
End If
rstDescription.Close

write_sp_log adoConn, 500, "spsp_GetStandingAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
" at contract_info_mfund.asp"


'Get Prima Bruta
Sql = "sppl_GetSumaAportes " & Contract
rstDescription.Open Sql, adoConn
If rstDescription.BOF And rstDescription.EOF Then
	PrimaBruta = 0
Else
	PrimaBruta = rstDescription.Fields(0)
End If
rstDescription.Close
write_sp_log adoConn, 500, "sppl_GetSumaAportes", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
" at contract_info_mfund.asp"


'Load AS400 Growth and Taxes


	''SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNIONÓ EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN
	''Set objSkMtrust = Server.CreateObject("Conexion.conectar")
	GrowthParams = Application("GrowthTaxes") & "?contract=" & Contract & "&date=" & InputDate

	Set objSkMtrust = CreateObject("MSXML2.XMLHTTP")
	objSkMtrust.Open "GET", GrowthParams , false
	objSkMtrust.setRequestHeader "Content-Type","text/xml"
	objSkMtrust.Send(null)
	GrowthParamsInfo = objSkMtrust.responseText


	''GrowthParamsInfo = objSkMtrust.urlReader(GrowthParams)


'Load AS400 Contract's Accounts


''Set objSkMtrust = Server.CreateObject("Conexion.conectar")


	Accounts = Application("Accounts") & "?contract=" & Contract
	objSkMtrust.Open "GET", Accounts , false
	objSkMtrust.Send(null)
	Accounts = objSkMtrust.responseText

''Accounts = objSkMtrust.urlReader(Accounts)
'Response.Write Accounts

Beneficiaries = Application("Beneficiaries") & "?contract=" & Contract
''Beneficiaries = objSkMtrust.urlReader(Beneficiaries)

	objSkMtrust.Open "GET", Beneficiaries , false
	objSkMtrust.Send(null)
	Beneficiaries = objSkMtrust.responseText


Terceros = Application("Terceros") & "?contract=" & Contract

''Terceros = objSkMtrust.urlReader(Terceros)

	objSkMtrust.Open "GET", Terceros , false
	objSkMtrust.Send(null)
	Terceros = objSkMtrust.responseText


	''FIN SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNIONÓ EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN


'Display info
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
						Response.Write "NdS"
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
						If rstDescription.BOF And rstDescription.EOF Then
							Response.Write arrContract(0,0)
						Else
							Response.Write rstDescription.Fields("Descripcion")
						End If
						CloseTd
						rstDescription.Close
					End If
					OpenTd "tbody", "width=50%"
						Response.Write "&nbsp;"
					CloseTd					
				CloseTr
				OpenTr ""
					OpenTd "tbody", ""
						'Response.Write "Afiliación" '<I&T - DMPC 2009/02/09 - Modificado por proceso de Referencia Unica - Recaudos>
						Response.Write "Contrato"
					CloseTd
					OpenTd "tbody", ""
						Response.Write ":"
					CloseTd					
					OpenTd "tbody", ""
						' Response.Write arrContract(1,0) '<I&T - DMPC 2009/02/09 - Modificado por proceso de Referencia Unica - Recaudos>
						response.Write Reference
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
						Response.Write "Fecha de ingreso"
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
				
				
				
			CloseTable
			'Values
			
			OpenTable "100%", ""
				
' Se reemplazan estas lineas por el llamado a una pagina en el servidor de Base de datos 400_files
' que contiene la informacion consolidada de todos saldos de Mtcor growth_Taxes.asp
' Rafael Lagos Noviembre 26 - 2002
'				OpenTr ""
'					OpenTd "tbody", "width=30%"
'						Response.Write "Retiros Realizados"
'					CloseTd
'					OpenTd "tbody", "width=5% align=left"
'						Response.Write ":"
'					CloseTd					
'					OpenTd "tbody", "align=right width=15%"
'
'						If IsNull(RetRealizados) Then
'							Response.Write "$0.00"
'						Else
'							Response.Write FormatCurrency(RetRealizados, 2)
'						End If
'					CloseTd
'					OpenTd "tbody", "width=50%"
'						Response.Write "&nbsp;"
'					CloseTd
'				CloseTr
'				OpenTr ""
'					OpenTd "tbody", "width=30%"
'						Response.Write "Impuestos Cobrados"
'					CloseTd
'					OpenTd "tbody", "width=5% align=left"
'						Response.Write ":"
'					CloseTd					
'					OpenTd "tbody", "align=right width=15%"
'
'						If IsNull(TaxCobrados) Then
'							Response.Write "$0.00"
'						Else
'							Response.Write FormatCurrency(TaxCobrados, 2)
'						End If
'					CloseTd
'					OpenTd "tbody", "width=50%"
'						Response.Write "&nbsp;"
'					CloseTd
'				CloseTr
				
				Response.Write GrowthParamsInfo
				
'				OpenTr ""
'					OpenTd "balance", ""
'						Response.Write "Saldo total"
'					CloseTd
'					OpenTd "balance", ""
'						Response.Write ":"
'					CloseTd					
'					OpenTd "balance", "align=right"
'						'If IsNull(arrContract(9,0)) Then
'						'	Response.Write FormatCurrency(0)
'						'Else
'						'	Response.Write FormatCurrency(arrContract(9,0), 2)
'						'End If
'						Response.Write Pagina						
'					CloseTd
'					OpenTd "tbody", ""
'						Response.Write "&nbsp;"
'					CloseTd
'				CloseTr

				OpenTr ""
					OpenTd "", ""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			CloseTable
'			<I&T> legal role

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
			
			OpenTable "100%", ""
				OpenTr ""
					OpenTd "", ""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead", ""
						Response.Write "Cuentas registradas para este contrato"
					CloseTd
				CloseTr
			CloseTable
				'Accounts Table --- START
				OpenTable "", "'' border=1"
					Response.Write Accounts
				CloseTable
				'Accounts Table --- END
			Response.Write "<br>"
			'Beneficiaries Table -- START
			OpenTable "", ""
				OpenTr ""
					OpenTd "thead", ""
						Response.Write "Beneficiarios registrados"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "", "'' border=1"
				Response.Write Beneficiaries
			CloseTable
			'Beneficiaries Table -- END
			Response.Write "<br>"
			'Beneficiaries Table -- START
			OpenTable "", ""
				OpenTr ""
					OpenTd "thead", ""
						Response.Write "Terceros registrados"
					CloseTd
				CloseTr
			CloseTable
			OpenTable "", "'' border=1"
				Response.Write Terceros
			CloseTable
			'Beneficiaries Table -- END
			Response.Write "<br>"
			'Build MultiFund Information
			'If RTrim(arrContract(0,0)) = "MFUND" Or  RTrim(arrContract(0,0)) = "MTCOR" Then
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
				write_sp_log adoConn, 500, "spem_GetAssetAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
				" at contract_info_mfund.asp"
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
					For I = 0 To UBound(arrAsset)-1
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
								write_sp_log adoConn, 500, "sppl_getvalorhistund", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & _
								" at contract_info_mfund.asp"
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
			'End If
				
		CloseTd
	CloseTr
	
	'=======================================================
    '======modificado R. lagos Febrero 2  2002
    '====== agregar validacion PAS
    '=======================================================
    
    if Request.Form("Pas")="S"    then      
      frase="Este contrato est&aacute; bajo el servicio Premium Advice Strategist; las transferencias entre fondos, " & _
				    "cambios de Standing Allocation y D.C.A. no est&aacute;n disponibles"
          openTr ""
            openTd "''",""
			  OpenTable "90%","border=1"
				OpenTr "class=teven"
					OpenTd "thead","align=center"
						Response.Write frase
					CloseTd
				CloseTr
			  CloseTable           
			 closetd
		   closetr	 
     end if
    '=======================================================
    '== fin modificacion
    '=======================================================
	
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
