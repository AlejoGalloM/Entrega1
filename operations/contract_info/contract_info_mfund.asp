<%@ Language=VBScript  %>
<%
'===================================================================================
'File Name:		contract_info_Mfund.asp 500
'Path:				contract_info/
'Created By:		Andres Felipe Orozco 2001/05/29
'Modified:	Andres Felipe Orozco  2001/06/19
'Last Modified:			
'						I&T - WTG 2007/12/18 Llamado a sistema de roles legales
'						M Cardozo 2007/10/24 Bandera para enviar a Tax. Subo la Alerta Fondos Cerrados
'						S Soriano 2007/08/13 Desplegar mensaje si un contrato no esta habilitado para retiros (PAC).
'						S Soriano Mayo 30 2007 Add omni source data
'						J Carreño Septiembre 11 2003 Modify get insurance information
'						Fabio Calvache Agosto 1 2003 Add DocType Marathon
'						R. Lagos 2002/06/19 remove Millas information 
'						G.Pinerez   2002/06/19 PAS Information long name fixing 1.1
'						G Pinerez  2002/05/30  PAS Information long name
'						R. Lagos  2002/02/05  Added Pas
'						A. Orozco  2001/10/05
'						A. Orozco 2001/10/09
'						Guillermo Aristizabal 2001/10/11
'						A. Orozco 2001/10/25
'						A. Orozco 2002/01/04 Accounts info direct from as400
'						javier VArgas 2005/30/06 add web services solution
'						I&T - WTG 20080313 Inclusion de consulta para clientes core
'						I&T-WTG 20080522 Formato $ para el saldo disponible
'						M Cardozo 2009/08/18 Bandera para enviar a Tax retiro express PC17-6152
'						A Figueroa	2010-06-02 Bandera para indicar si tiene garantia de credito HSBC para TAX
'                      				 I&T - Nelson D. Peña N. 2010/10/12 Incluir perfil de riesgo del contrato
'                      				 I&T - Nelson D. Peña N. 2010/11/25 Incluir perfil real de riesgo del contrato
'                      				 I&T - Camilo Gutierrez. 2011/02/10 Cambios perfil de riesgo del contrato
'                      				 I&T - Nelson D. Peña N. 2011/03/24 Incluir validación de las variables para los contratos de planes corporativos
'					   Oscar Diaz 2012/11/06
'					   Se utiliza una función para comprobar los planes corporativos
'					   y evitar el uso de variables quemadas en el asp
'                      I&T Daniel Rodriguez 2015-11-23 Se elimina código de Insurance no implementado ISSUE CO01-88317 
'                      I&T Daniel Rodriguez 2016-05-06 Se realiza validación disponibilidad de servicio OMNI GEMINI 100782
'Parameters:		User must be logged on
'						Session("docNumber")
'Returns:			MFUND contract information
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
<!--#include file="../_pipeline_scripts/mfundCorporativoScripts.asp"-->
<%
Authorize 5,1
Dim Sql 'SQL Sentences holder
Dim adoConn 'Database Connection
Dim rstDescription 'Results recordset
Dim objRst	' Strategist PAS recordset
Dim rs, cn
Dim arrContract
Dim arrAsset, arrAccounts
Dim arrStanding, Status
Dim PrimaBruta
Dim objSkMtrust, Accounts
Dim I,J 'Asset & Standing Alloc counters
Dim Contract, Product, Plan, ClientId, DocType, Name, Phone, OU, Pas, PasName, Frase 
Dim IdObjetivo ' javier  yaxa
Dim Beneficiaries
Dim Programa '' Jimmy ospino 
Dim arrprogram '' Jimmy ospino 
Dim PartID,SourceID,SourceName,CorpProduct,PermiteRetiros,SaldoMinRetiro,MontoMinRetiro, NombreEmpresa 'Sonia Soriano
Dim arrfcerrados, rstfcerrados, fcerrados 'MCardozo
Dim IsLegalRole ' I&T - WTG
Dim Reference '<I&T - DMPC>
Dim arrretiroexpress, rstretiroexpress, retiroexpress 'MCardozo
Dim garantiahsbc 'JFigueroa
Dim riskProfile 'I&T - Nelson D. Peña N.
Dim realRiskProfile 'I&T - Nelson D. Peña N.
Dim getContractInfo 'I&T - Nelson D. Peña N. - 2011/03/24 Bandera para validar si se debe consultar los datos de los contratos de planes corporativos
    getContractInfo = False
Dim valueITP, idType
'Jose Alejandro Figueroa - Proyecto Advice Tools - 2011-03-16
Dim hasContractInfo
    hasContractInfo = False
DIM objConn

Session.contents("errorOMNI") = "N"

'====================
' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
'====================
Dim IsCoreDescription, arrComplement
	IsCoreDescription = "No se Conoce"

	IsLegalRole = False
	Pas = Request.Form("Pas")
	PasName = Request.Form("PasName")
	Contract = Request.Form("Contract")
	Product = Request.Form("Product")
	Plan = Request.Form("Plan")
	Session("name")=ClientId
	ClientId = Request.Form("ClientId")
	DocType = Request.Form("DocType")
	Name = Request.Form("Name")

	Phone = Request.Form("Phone")
	OU = Request.Form("OU")

write_dataLog Response.Status,"contract_info_mfund.asp","process_selection Loaded by: " & Session("sp_miLogin") & " selected  for the Contract " & Contract,ClientId,"ContratoRetirosPAC_GetForContrato " & Contract & " - ContratoRetirosPAC_GetForContrato " & Contract,"N/A","null","Consulta","N/A"


	Session("seleccionContrato")="false"

	PartID = Request.Form("PartID")
	SourceID = Request.Form("SourceID")
	SourceName = Request.Form("SourceName")
	CorpProduct = Request.Form("CorpProduct")
	PermiteRetiros = Request.Form("PermiteRetiros")
	SaldoMinRetiro = Request.Form("SaldoMinRetiro")
	MontoMinRetiro = Request.Form("MontoMinRetiro")
	NombreEmpresa = Request.Form("NombreEmpresa")

	'Alejandro Figueroa - Proyecto Advice Tools - 2011-03-16
	'La funcionalidad aplica para todos los contratos Multifund Individual y para los Corporativo tipo CAHC (los demás no)
	' Oscar Diaz 2012-10-31 Se modifica para que no existan valores quemados de los planes
	If trim(Product) = "MFUND" And VerificarPlanProducto(Application("CorpPlanesAll"),trim(Plan)) <>  true  Then
		hasContractInfo = True
	End If

	'I&T - Nelson D. Peña N. - 2011/03/24 
	' Si el plan del contrato es un plan corporativo y las variables PermiteRetiros, SaldoMinRetiro y MontoMinRetiro son nulas o vacias se activa la bandera getContractInfo para obtener los datos del contrato.
	' Oscar Diaz 2012-10-31 Se modifica para que no existan valores quemados de los planes
	If trim(Product) = "MFUND" And ( VerificarPlanProducto(Application("CorpPlanes"),trim(Plan)) = true) And (isnull(PermiteRetiros) Or PermiteRetiros = "" Or isnull(SaldoMinRetiro) Or SaldoMinRetiro = "" Or isnull(MontoMinRetiro) Or MontoMinRetiro = "" ) Then		
		getContractInfo = True
	End If

	If (Session.Contents("SiteRetuns") = "2") Or getContractInfo Then 
		Dim auxArrDatosSource, auxDatosSource, datosSource ,arrDatosSource,planID, PermiteRetirosOmni ,strSql ,arrDatosEmpresa 
		'===================Obtener solo datos de la source del contrato equivalente - S.Soriano 2007/05/30============
		' Oscar Diaz 2012-10-31 Se modifica para que no existan valores quemados de los planes
		If VerificarPlanProducto(Application("CorpPlanes"),trim(Plan)) = true Then
			'Response.Write " -> Consultando información contrato PAC"
			Set objConn = GetConnPipelineDB

			datosSource = GetSourceInfo(trim(Contract))
            'Validación de respuesta de servicio OMNI
            if trim(datosSource) <> "ERROR TRAYENDO DATOS SOURCE"  then	
			    arrDatosSource = split(datosSource, "-")
			    partID = trim(arrDatosSource(0))
			    planID = trim(arrDatosSource(1))
			    SourceID = trim(arrDatosSource(2))
			    SourceName = trim(arrDatosSource(3))
			    CorpProduct = trim(arrDatosSource(4))
			    PermiteRetirosOmni = trim(arrDatosSource(5))
			    MontoMinRetiro = trim(arrDatosSource(6))
			    SaldoMinRetiro = trim(arrDatosSource(7))	
            else 
		        PermiteRetirosOmni = "N"
		        Session.contents("errorOMNI") = "Y"
	        end if

			'==Inicio Traer de la tabla temporal de contratosRetirosPAC el flag correspondiente al retiro habilitado, si existe en la tabla====
			If PermiteRetirosOmni = "N" Then
				strSql = "ContratoRetirosPAC_GetForContrato " & Contract
				Set objRst = Server.CreateObject("ADODB.Recordset")
				objRst.Open strSql, objConn

				If objRst.BOF And objRst.EOF Then
					PermiteRetiros = PermiteRetirosOmni
				Else
					arrContratoPAC = objRst.GetRows()
					If arrContratoPAC(11,0) = True Then
						PermiteRetiros = "Y"
					Else
						If arrContratoPAC(11,0) = False Then
							PermiteRetiros = "N"
						Else
							PermiteRetiros = PermiteRetirosOmni
						End If
					End If
				End If
				objRst.Close
			Else
				PermiteRetiros = PermiteRetirosOmni
			End If

			'====Fin Traer datos tabla temporal====

			'****** Inicio Traer Nombre Empresa para contrato PAC
			strSql = "ContratosPAC_InfoEmpresa " & Contract
			Set objRst = nothing
			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open strSql, objConn

			If objRst.BOF And objRst.EOF Then
				nombreEmpresa = "NA"
			Else
				arrDatosEmpresa = objRst.GetRows()
				nombreEmpresa = arrDatosEmpresa(5,0)
			End If

			objRst.Close
			SET objConn= NOTHING

			'****** Fin Traer Nombre Empresa ========

			Session.Contents("PermiteRetiros") = PermiteRetiros
			Session.Contents("MontoMinRetiro") = MontoMinRetiro
			Session.Contents("SaldoMinRetiro") = SaldoMinRetiro
			Session.Contents("NombreEmpresa") = nombreEmpresa
		Else 
			datosSource = ""
		End If
		'===================Fin Obtener solo datos de la source del contrato equivalente - S.Soriano 2007/05/30=========
	End If 
    
	Set adoConn = GetConnpipelineDB

	'====================
	' <I&T - WTG: ISCORE (20080313) inclusion de consulta si un cliente es Core>
	'====================
	write_sp_log adoConn, 500, "Iscore : " & CStr(ClientId) + ":" & DocType, Contract, Product, Plan, ClientId, 0, "", "contract_info_mfund.asp " & "Loaded by " & Session("sp_miLogin")
	
	Set rstDescription = Server.CreateObject("ADODB.Recordset")

	If ClientId = "" Then 
        ClientId = 0
    End If

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
	write_sp_log adoConn, 500, "IdObjetivo : " + Request.Form("idObjetivo"), Contract, Product, Plan, ClientId, 0, "", "contract_info_mfund.asp " & "Loaded by " & Session("sp_miLogin")

	'==========================================================================
	''<I&T - DMPC 2009/02/13 - Modificado por proceso de Referencia Unica - Recaudos>
	' Obtiene la referencia única a partir del Contrato y el producto
	'==========================================================================
	 Reference = GetReferenciaUnica( Product, Contract)
	'==========================================================================

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

	Sql = "sppl_GetDetallesContrato " & Contract & ", '" & Product & "', '" & Plan & "'"
	rstDescription.Open Sql, adoConn

	If rstDescription.BOF And rstDescription.EOF Then
		arrContract = 0
	Else
		arrContract = rstDescription.GetRows()
	End If

	rstDescription.Close

	If IsArray(arrContract) Then
		If (ISNULL(arrContract(107,0))) Then
            Sql = "spem_GetDescriptionProgram '" & plan & "'"
			rstDescription.Open Sql, adoConn

			If rstDescription.BOF And rstDescription.EOF Then
				arrprogram = 0
			Else
				arrprogram = rstDescription.GetRows()
				Programa = arrprogram(0,0)
			End If

			rstDescription.Close
		Else
			Programa = arrContract(107,0)
		End If
	End If

	'El contrato tiene Fondos Cerrados M.Cardozo 2007/10/24
	set rstfcerrados = Server.CreateObject("ADODB.Recordset")
	Sql = "sppl_Contrato_GetTieneAlgunFondoCerradoVenta '" & Product & "','" & Contract & "'"
	rstfcerrados.Open Sql, adoConn
	fcerrados = "N"

	If rstfcerrados.BOF And rstfcerrados.EOF Then
        '
	Else
		If (rstfcerrados.Fields("TieneFondosCerradosVenta")="1") Then
			fcerrados="S"
		End If
	End If

	rstfcerrados.Close
	session.Contents("FCerrados")=fcerrados

	'Todos los productos de MFUND tienen disponible retiro Experess AJA
	session.Contents("RetiroExpress") = "S"

	'============================================================================
	''<Alejandro Figueroa 2010-06-02 - Verifica si tiene garantia de credito HSBC
	'============================================================================
	garantiahsbc = False
	Dim hsbc_xmlDoc, hsbc_HTTP
		
	Set hsbc_HTTP = CreateObject("Microsoft.XMLHTTP")
	Set hsbc_xmlDoc = CreateObject("MSXML.DOMDocument")
	hsbc_xmlDoc.Async = False

	' Invoca el ws de disponible
	hsbc_HTTP.Open "GET", Application("WSAvailableService") & "/javaGet/GetHsbcGuaranteeValue?contract=" & cstr(Contract) & "&product=" & Product, False
	hsbc_HTTP.Send(null)

	If hsbc_xmlDoc.load(hsbc_HTTP.responseXML) Then
		If hsbc_xmlDOC.documentElement.text <> "0" Then
			garantiahsbc = True
		End If
	End If

	Dim Pas_ 
	Pas_ = "N"
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "exec sppl_getpas '" & product & "'," & contract
	rs.Open sql, adoConn

	If Not rs.EOF And Not rs.BOF Then
		If rs(0) = 1 Then			
			Pas_ ="S"
		End If
	End If

	If IsArray(arrContract) Then
		If Not (ISNULL(arrContract(15,0)) Or (ISNULL(arrContract(35,0)))) Then
            AuthorizeContractAccess Cstr(Session("sp_AccessLevel")), CStr(Session("sp_IdAgte")), CStr(arrContract(15,0)), CStr(arrContract(35,0))
		End If
	End If

	write_sp_log adoConn, 500, "sppl_GetDetallesContrato", Contract, Product, Plan, ClientId, 0, "", "contract_info_mfund.asp " & "Loaded by " & Session("sp_miLogin")

	If IsArray(arrContract) Then
		'Get Status Description
		Sql = "spsp_GetStatusDescription '" & Product & "', '" & arrContract(14,0) & "'"
		rstDescription.Open Sql, adoConn

		If rstDescription.BOF And rstDescription.EOF Then
			Status = arrContract(14,0)
		Else
			Status = rstDescription.Fields("descripcion")
		End If

		rstDescription.Close
		write_sp_log adoConn, 500, "spsp_GetStatusDescription", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_mfund.asp"
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
	write_sp_log adoConn, 500, "sppl_CuentasRegXContrato", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_mfund.asp"

	Sql = "spsp_GetStandingAllocation " & Contract & ", '" & Product & "'"
	rstDescription.Open Sql, adoConn

	If rstDescription.BOF And rstDescription.EOF Then
		arrStanding = 0
	Else
		arrStanding = rstDescription.GetRows()
	End If
    
	rstDescription.Close
	write_sp_log adoConn, 500, "spsp_GetStandingAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_mfund.asp"

	Sql = "sppl_GetSumaAportes " & Contract
	rstDescription.Open Sql, adoConn

	If rstDescription.BOF And rstDescription.EOF Then
		PrimaBruta = 0
	Else
		PrimaBruta = rstDescription.Fields(0)
	End If

	rstDescription.Close
	write_sp_log adoConn, 500, "sppl_GetSumaAportes", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_mfund.asp"
    
	Set objRst = Server.CreateObject("ADODB.Recordset")		
	Sql = "exec spsp_getPasName '" & Product & "'," & Contract
	objRst.Open Sql, adoConn

	If Not objRst.BOF And Not objRst.EOF Then
	   Pas = "S"
	   PasName = objRst(0)
	Else
	   Pas = "N"   
	   PasName = ""			   
	End If

	objRst.Close

	'==========================================================================
	'==end info insurance
	'==========================================================================

	If hasContractInfo Then
		Sql = "Relacionamiento..Contract_Profile_GetProfileByContractByCountry '" & CStr(DocType) & CStr(ClientId) & "', 'CO', " & RTrim(Contract) & ", '" & CStr(Product) & "'"
		rstDescription.Open Sql, adoConn

		If rstDescription.BOF And rstDescription.EOF Then
			riskProfile = ""
		Else	
			riskProfile = rstDescription.GetRows()
		End If

		rstDescription.Close

		Sql = "Relacionamiento..Contract_Profile_GetRealRiskProfileByContractByCountry '" & CStr(DocType) & CStr(ClientId) & "', 'CO', " & RTrim(Contract) & ", '" & CStr(Product) & "'"
		rstDescription.Open Sql, adoConn

		If rstDescription.BOF And rstDescription.EOF Then
			realRiskProfile = ""
		Else	
			realRiskProfile = rstDescription.GetRows()
		End If

		rstDescription.Close

		'==========================================================================
		'==END Perfil real de riesgo
		'==========================================================================
	End If
	'=================================================================
	'fin borrado parcialmente nuevo reglamento JM 
	'=================================================================
    		
	OpenHTML
		OpenHead
			PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
			PlaceMeta "Pragma", "", "no_cache"
			PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	%>
	<link href="/cake/asp.css" type="text/css" rel="stylesheet" />
	<style>
	a.info{
		position:relative; /*this is the key*/
		z-index:24; background-color:#fff;
		color:#FF0000;
		text-decoration:none}

	a.info:hover{z-index:25; background-color:#DDD}

	a.info span{display: none}

	a.info:hover span{ /*the span will display just on :hover state*/
		display:block;
		position:absolute;
		top:1em; left:-5em; width:15em;
		border:5px solid #006f53;
		background-color:#DDD; color:#000;
		text-align: center}		
	</style>
	<%
		CloseHead
		OpenBody "", ""
		If InStr(1, Request.ServerVariables("HTTP_REFERER"), "menu.asp") = 0 Then
			'Reload Left Menu -- START
			OpenForm "menu_left", "post", "../menu/menu.asp", "target=menu"
				Session.Contents("ClientId") = ClientId
				PlaceInput "Name", "hidden", Request.Form("Name"), ""
				PlaceInput "ClientId", "hidden", ClientId, ""
				PlaceInput "DocType",  "hidden", DocType, ""
				PlaceInput "Phone", "hidden", Phone, ""
				PlaceInput "Contract", "hidden", Contract, ""
				PlaceInput "Product", "hidden", Product, ""
				PlaceInput "Plan", "hidden", Plan, ""
				PlaceInput "OU", "hidden", OU, ""
				PlaceInput "Pas", "hidden", Pas_, ""
				'===  Inicio S.Soriano variables retiros pac ========
				PlaceInput "PermiteRetiros", "hidden", PermiteRetiros, ""
				PlaceInput "MontoMinRetiro", "hidden", MontoMinRetiro, ""
				PlaceInput "SaldoMinRetiro", "hidden", SaldoMinRetiro, ""
				'==== Fin S.Soriano variables retiros pac ==========
				
				'== Alejandro Figueroa 2010-06-02
				'' Si tiene garanta HSBC se comporta igual que si tuviera fondos cerrados - No pemite retiros totales en TAX
				If garantiahsbc Then
					PlaceInput "FCerrados", "hidden", "S", ""
				Else
					PlaceInput "FCerrados", "hidden", fcerrados, ""
				End If

				PlaceInput "RetiroExpress", "hidden", retiroexpress, ""
			CloseForm
	%>
	<script language="javascript">
				document.menu_left.submit();
	</script>
	<%
			'Reload Left Menu -- END
		End If
	
		If IsArray(arrContract) Then  'Detalle del contrato
			Response.Write "<br>"
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
							CloseTd
						CloseTr
					CloseTable
					Response.Write "<br>"
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
							' Oscar Diaz 2012-10-31 Se modifica para que no existan valores quemados de los planes
								If VerificarPlanProducto(Application("CorpPlanes"), Trim(Plan)) = TRUE Then
									Response.Write "Mfund - " & CorpProduct
									CloseTd	
								Else
									Sql = "spem_GetDescriptionProduct " & arrContract(0,0)
									rstDescription.Open Sql, adoConn
										Response.Write rstDescription.Fields("Descripcion")
									CloseTd
									rstDescription.Close
								End If
							End If
							OpenTd "tbody", "width=50%"
								Response.Write "&nbsp;"
							CloseTd					
						CloseTr
						'======================================================
						'Camilo Gutierrez I&T 10-02-2011 Agregar Nombre de contrato
						'======================================================
						If hasContractInfo Then 
							OpenTr ""
								OpenTd "tbody", ""
									Response.Write "Nombre Contrato"
								CloseTd
								OpenTd "tbody", ""
									Response.Write ":"
								CloseTd					
								OpenTd "tbody", ""							
									If arrContract(103,0) = "" OR isnull(arrContract(103,0)) Then							
										Response.Write "No definido <img title='Su cliente podrá definir esta información a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' /><br/>"
									Else
										Response.Write arrContract(103,0)								
									End If
								CloseTd
								OpenTd "tbody", ""
									Response.Write "&nbsp;"
								CloseTd					
							CloseTr
						End If
						'======================================================
						'End Camilo Gutierrez I&T 10-02-2011 Agregar Nombre de contrato
						'======================================================
						OpenTr ""
							OpenTd "tbody", ""
								Response.Write "Número Contrato"
							CloseTd
							OpenTd "tbody", ""
								Response.Write ":"
							CloseTd					
							OpenTd "tbody", ""
								response.Write Reference
							CloseTd
							OpenTd "tbody", ""
								Response.Write "&nbsp;"
							CloseTd					
						CloseTr
						'======================================================
						'Camilo Gutierrez I&T 10-02-2011 Agregar Objetivo de contrato
						'======================================================
						If hasContractInfo Then 
							OpenTr ""
								OpenTd "tbody", ""
									Response.Write "Objetivo Contrato "
								CloseTd
								OpenTd "tbody", ""
									Response.Write ":"
								CloseTd
								OpenTd "tbody", ""							
									If arrContract(104,0) = "" Or ISNULL(arrContract(104,0)) Then							
										Response.Write "No definido <img title='Su cliente podrá definir esta información a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' /><br/>"
									Else
										Response.Write arrContract(104,0)								
									End If
								CloseTd
								OpenTd "tbody", ""
									Response.Write "&nbsp;"
								CloseTd
							CloseTr
						End If
						'======================================================
						'End Camilo Gutierrez I&T 10-02-2011 Agregar Objetivo de contrato
						'======================================================
						' Oscar Diaz 2012-10-31 Se modifica para que no existan valores quemados de los planes
						If VerificarPlanProducto(Application("CorpPlanes"),trim(Plan)) = True Then							
							OpenTr ""
								OpenTd "tbody", ""
									Response.Write "Cuenta Individual"
								CloseTd
								OpenTd "tbody", ""
									Response.Write ":"
								CloseTd					
								OpenTd "tbody", ""
									Response.Write PartID
								CloseTd
								OpenTd "tbody", ""
									Response.Write "&nbsp;"
								CloseTd					
							CloseTr				
							OpenTr ""
								OpenTd "tbody", ""
									Response.Write "Subcuenta"
								CloseTd
								OpenTd "tbody", ""
									Response.Write ":"
								CloseTd					
								OpenTd "tbody", ""
									Response.Write SourceID & " - " & SourceName
								CloseTd
								OpenTd "tbody", ""
									Response.Write "&nbsp;"
								CloseTd					
							CloseTr
							OpenTr ""
								OpenTd "tbody", ""
									Response.Write "Empresa"
								CloseTd
								OpenTd "tbody", ""
									Response.Write ":"
								CloseTd					
								OpenTd "tbody", ""
									Response.Write NombreEmpresa
								CloseTd
								OpenTd "tbody", ""
									Response.Write "&nbsp;"
								CloseTd					
							CloseTr
						End If
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
								Response.Write "Programa"
							CloseTd
							OpenTd "tbody", ""
								Response.Write ":"
							CloseTd					
							OpenTd "tbody", ""
							select case trim(Plan)
								Case "TL01"
								   Response.Write "Total Life"
								Case "FV01"
									Response.Write "FonVida"   
								Case "FB01"
									Response.Write "Fibac"   
								Case "IG01"
									Response.Write "Inversion Gold"   
								Case Else
								 Response.Write Programa
							End select 
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
			
						SELECT CASE trim(ucase(arrContract(0,0)))
							CASE "MFUND"				
								'========================================================================================
								'begin insurance information april 26 2002 add jecv
								'========================================================================================
								'--------------------------- MENCO: 01-06-2012, Mostrar la suma asegurada para los contratos SKCS --------- INICIO
								If TRIM(Plan) = "SKCS" Then
									Dim valoramparo
									valoramparo = arrContract(106,0)
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Suma Asegurada Amparo por Fallecimiento"
										CloseTd
										OpenTd "tbody", ""
										Response.Write ":"
										CloseTd
										OpenTd "tbody", ""
											IF arrContract(106,0) = "" Or ISNULL(arrContract(106,0)) Then
												Response.Write "N/A"
											Else
												Response.Write FormatCurrency(arrContract(106,0),2) 'Suma Asegurada en sigscg..poliza
											End If						
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd										
									CloseTr
								'----------------------------------------------------------------------------------------------------------------------
									valueITP = 0
									idType = 0
									Select Case arrcontract(4, 0)
									Case "Cédula de Ciudadanía"
										idType = 1
									Case "Pasaporte"
										idType = 0
									Case "Cédula de Extranjeria"
										idType = 4
									Case "Identificación Tributaria"
										idType = 2
									Case Else
										idType = 1
									End Select
									
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Suma Asegurada Amparo por ITP"
										CloseTd
										OpenTd "tbody", ""
										Response.Write ":"
										CloseTd
										OpenTd "tbody", ""
											IF ISNULL(valoramparo) Then
												Response.Write "N/A"
											Else
												Sql = "DECLARE	@return_value int EXEC	@return_value = Pharos..GetITPValue " & idType & "," & valoramparo & ", '" & ClientId & "' SELECT	'Return Value' = @return_value"
												rstDescription.Open Sql, adoConn
												Response.Write FormatCurrency(rstDescription.Fields("Return Value"),2)
												CloseTd
												rstDescription.Close
												'response.write Sql
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd										
									CloseTr
								'----------------------------------------------------------------------------------------------------------------------
								End If
								'--------------------------- MENCO: 01-06-2012, Mostrar la suma asegurada para los contratos SKCS --------- FIN
								OpenTr ""
									OpenTd "tbody", ""
										Response.Write "Portafolio de Inversión"
									CloseTd
									OpenTd "tbody", ""
										Response.Write ":"
									CloseTd
									OpenTd "tbody", ""
										If Pas_="S" Then      
											Response.Write "Strategist"
										Else
											Response.Write "Individual"
										End If
										'Cambio M.Cardozo subo el warning fondos cerrados"	
										If fcerrados = "S" Then
											Response.Write "<font color=red><strong> -P. Cerrados- </strong></font>"
											'MENCO: 03-10-2012, por solicitud de A. Jaramillo/Juan P. Maya
											'Response.Write "<font color=gray><strong> -P. Cerrados- </strong></font>"
										End If
									CloseTd
									OpenTd "tbody", ""
										Response.Write "&nbsp;"
									CloseTd										
								CloseTr
								'==========================================================================
								' Incluir perfil de riesgo del contrato 2010/10/12 I&T - Nelson D. Peña N.
								'==========================================================================	
								If hasContractInfo Then 
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Perfil de inversión del contrato"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd
										OpenTd "tbody", ""
											If IsArray(riskProfile) Then
												Response.Write riskProfile(2,0)
											Else
												Response.Write "No definido <img title='Su cliente podrá definir esta información a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' /><br/>"
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd										
									CloseTr
									'==========================================================================
									' Incluir perfil real de riesgo del contrato 2010/11/25 I&T - Nelson D. Peña N.
									'==========================================================================	
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Perfil real del contrato"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd
										OpenTd "tbody", ""
											If IsArray(realRiskProfile) Then
												Response.Write realRiskProfile(2,0)
											Else
												Response.Write "El contrato no tiene asignado un perfil real."
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd										
									CloseTr
								End If
								
							CASE "SIPEN"
								pas="S"
						END SELECT 
						'Alejandro Figueroa - Si tiene garantias HSBC se escribe texto oculto como verificacion"	
						If garantiahsbc Then
							OpenTr ""
								OpenTd "tbody", "colspan=4"
									Response.Write "<font color=white><strong>Contrato con saldo en garantía de crédito con HSBC</strong></font>"
								CloseTd
							CloseTr
						End If
					CloseTable

					'--------------------- MENCO: 13-06-2012, cambio para obtener los datos desde un web service --------------- INICIO
					If arrContract(8,0) = "" OR IsNull(arrContract(8,0)) Then 'MENCO: 31-08-2012, Optimización de tiempo de carga de la página (Contrato.FecTerminacion)					
						'Contrato valido
						If arrContract(9,0) = "" Or IsNull(arrContract(9,0)) Then
							'Saldo nulo
							Response.Write "<br>"
							OpenTable "90%","border=1"
								OpenTr "class=teven"
									OpenTd "thead","align=center"
										Response.Write "No existe información de saldos desde TAX"
									CloseTd
								CloseTr
							CloseTable
						Else
							'[En desarrollo http://cobodvap01/SkCo.TaxBFacade/TaxFacade.asmx]							
							Dim strTaxSaldos, arrTaxSaldos
							strTaxSaldos = ""						
							'strTaxSaldos = InvokeTaxFacade("GetSaldosTB?strProducto=" & Product & "&intContrato=" & Contract)
							'MENCO: 05-09-2012, Ajuste por optimización. A. Jaramillo [se envía Producto, Contrato y Saldo]
							strTaxSaldos = InvokeTaxFacade("GetSaldosTax?strProducto=" & TRIM(Product) & "&intContrato=" & TRIM(Contract) & "&decSaldo=" & arrContract(9,0))							
							arrTaxSaldos = Split(strTaxSaldos) 'Se asume que el carácter delimitador es el carácter de espacio [" "]

							If IsArray(arrTaxSaldos) And LEN(strTaxSaldos) > 0 Then
								Response.Write "<br>"						
								OpenTable "100%", ""
									OpenTr ""
										OpenTd "thead", ""
												Response.Write "Valores"
										CloseTD
									CloseTR
								CloseTable
								OpenTable "100%", ""
									OpenTr ""
										OpenTd "tbody", "width=30%"
											Response.Write "Saldo capital"
										CloseTd
										OpenTd "tbody", "width=5%"
											Response.Write ":"
										CloseTd					
										OpenTd "tbody", "align=right width=15%"
											If IsNull(arrTaxSaldos(2)) Then
												Response.Write "$0.00"
											Else
												Response.Write FormatCurrency(arrTaxSaldos(2),2)
											End If
										CloseTd
										OpenTd "tbody", "width=50%"
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Saldo rendimientos"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd					
										OpenTd "tbody", "align=right"
											If IsNull(arrTaxSaldos(0)) Then
												Response.Write "$0.00"
											Else
												Response.Write FormatCurrency(arrTaxSaldos(0),2)
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Saldo total"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd					
											OpenTd "tbody", "align=right"
											If IsNull(arrTaxSaldos(1)) Then
												Response.Write "$0.00"
											Else
												Response.Write FormatCurrency(arrTaxSaldos(1),2)
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Cuenta contingente"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd					
										OpenTd "tbody", "align=right"
											If IsNull(arrTaxSaldos(3)) Then
												Response.Write "$0.00"
											Else
												Response.Write FormatCurrency(arrTaxSaldos(3),2)
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
									'------ 'MENCO: 12-06-2012
									If TRIM(Plan) = "SKCS" Then
										OpenTr ""
											OpenTd "tbody", ""
												Response.Write "Valor beneficio de permanencia*"
											CloseTd
											OpenTd "tbody", ""
												Response.Write ":"
											CloseTd					
											OpenTd "tbody", "align=right"
												If IsNull(arrTaxSaldos(7)) Then
													Response.Write "$0.00"
												Else
													Response.Write FormatCurrency(arrTaxSaldos(7),2)
												End If
											CloseTd
											OpenTd "tbody", ""
												Response.Write "&nbsp;"
											CloseTd
										CloseTr
									End If
									'------
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Total capital retirado"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd					
										OpenTd "tbody", "align=right"
											If IsNull(arrTaxSaldos(5)) Then
												Response.Write "$0.00"
											Else
												Response.Write FormatCurrency(arrTaxSaldos(5),2)
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Total rendimientos retirados"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd					
										OpenTd "tbody", "align=right"
											If IsNull(arrTaxSaldos(6)) Then
												Response.Write "$0.00"
											Else
												Response.Write FormatCurrency(arrTaxSaldos(6),2)
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
									OpenTr ""
										OpenTd "tbody", ""
											Response.Write "Primas brutas"
										CloseTd
										OpenTd "tbody", ""
											Response.Write ":"
										CloseTd					
										OpenTd "tbody", "align=right"
											If IsNull(arrTaxSaldos(4)) Then
												Response.Write "$0.00"
											Else
												Response.Write FormatCurrency(arrTaxSaldos(4),2)
											End If
										CloseTd
										OpenTd "tbody", ""
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
									'---
									OpenTr ""
										OpenTd "", ""
											Response.Write "&nbsp;"
										CloseTd
									CloseTr
								CloseTable
								
								If TRIM(Plan) = "SKCS" Then
									OpenTable "100%", ""
										OpenTr ""
											OpenTd "tbody", "colspan=3 valign=top"
												Response.Write "*Se hará efectivo una vez los aportes realizados en el Programa de Inversión cumplan 3 años de permanencia." 'MENCO: 12-06-2012
											CloseTd
										CloseTr
									CloseTable
								End If
							
							Else 'No existe informacion de saldos desde TAX
								Response.Write "<br>"
								OpenTable "90%","border=1"
									OpenTr "class=teven"
										OpenTd "thead","align=center"
											Response.Write "No existe información de saldos desde TAX"
										CloseTd
									CloseTr
								CloseTable
							End If
							
						End If
										
					Else 'MENCO: 31-08-2012, Ajustes para optimización de la página: Contrato.FecTerminacion es diferente de NULL
						Response.Write "<br>"
						OpenTable "90%","border=1"
							OpenTr "class=teven"
								OpenTd "thead","align=center"
									Response.Write "No existe información de saldos desde TAX"
								CloseTd
							CloseTr
						CloseTable
					End If

					'--------------------- MENCO: 13-06-2012, cambio para obtener los datos desde un web service --------------- FIN
				
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

					'=================================================
					'  CAMBIO PARA LLAMAR AL WEBSERVICE AS400 2007/08/27 APC 
					'=================================================
					Dim xmlDOC
					Dim bOK
					Dim HTTP
					Dim accion
					Dim valor
					Set HTTP = CreateObject("MSXML2.ServerXMLHTTP")
					Set xmlDOC =CreateObject("MSXML.DOMDocument")
					xmlDOC.Async=False
					accion = Application.Contents("AS400WS") & "/GetAccount"
					HTTP.Open "POST",accion,False
					HTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
					valor = "contract="& Contract
															
					HTTP.Send valor
					bOK = xmlDOC.load(HTTP.responseXML)
					OpenTable "90%", "'' border=0"
						OpenTr ""
							OpenTd "thead", ""
								If Not bOK Then
									Response.write "<br>ERROR TRAYENDO CUENTAS DEL CLIENTE </br>"	
								Else
									Dim objNodeListUrl, result
									result=""
									Set objNodeListUrl = xmlDOC.documentElement.selectNodes("//string")
									response.write objNodeListUrl.Item(0).Text
								End If
							CloseTd
						CloseTr
					CloseTable

					'=================================================
					'  FIN CAMBIO PARA LLAMAR AL WEBSERVICE AS400 2007/08/27
					'=================================================
                		
					'Build MultiFund Information
					If RTrim(arrContract(0,0)) = "MFUND" Then
						Response.Write "<br>"
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
						write_sp_log adoConn, 500, "spem_GetAssetAllocation", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_mfund.asp"
						
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
		
								For J = 0 To Ubound(arrAsset, 2) 'Rows [Ubound devuelve el mayor subíndice disponible para la dimensión indicada de una matriz]
									If (J Mod 2) = 0 Then
										OpenTr "class=todd"
									Else
										OpenTr "class=teven"
									End If								
									
									For I = 0 To UBound(arrAsset)
										If I <> 5 And I <> 6 Then
											If I <> 3 Then
												OpenTd "tbody", "align=Center"
											Else
												OpenTd "tbody", "align=Right"
											End If
										End If
										Select Case I
											Case 1
												Response.Write FormatNumber(arrAsset(I,J),2)
												
											Case 3	'Valor unidad ($)
												If IsNull(arrContract(11,0)) Then
													Sql = "sppl_getvalorhistund '" & FormatDateTime(Now(), 2) & "','" & FormatDateTime(Now(), 2) & "','" & arrAsset(I+2,J) & "'"
												Else
													Sql = "sppl_getvalorhistund '" & FormatDateTime(arrContract(11,0), 2) & "','" & FormatDateTime(arrContract(11,0),2) & "','" & arrAsset(I+2,J) & "'"
												End If
												
												rstDescription.Open Sql,adoConn
												
												If rstDescription.BOF And rstDescription.EOF Then
													Response.Write "<p align=center>N/A</p>"
												Else
													Response.Write FormatCurrency(rstDescription.Fields("valorunidad"), 6)
												End If
												
												rstDescription.Close
												
												write_sp_log adoConn, 500, "sppl_getvalorhistund", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_mfund.asp"
												
											Case 7
												If arrAsset(I,J)= False Then
													'MENCO: 03-10-2012, por solicitud de A. Jaramillo/Juan P. Maya (se modifica el estilo 'info' de la página)
													'Response.Write "<p align=center><font color=red><a class=info href='#'>NO<span><strong>PRECAUCIÓN</strong>:Este fondo está cerrado</span></a></font></p>"																										
													Response.Write "NO"
												Else
													Response.Write "SI"
												End If
												
											Case 5
												Response.Write ""
												
											Case 6
												Response.Write ""
        											
											Case Else
												If arrAsset(7,J)= False Then
													'MENCO: 03-10-2012, por solicitud de A. Jaramillo/Juan P. Maya (se modifica el estilo 'info' de la página)
													Response.Write "<p align=center><font color=red><a class=info href='#'>" & arrAsset(I,J) & "<span><strong>PRECAUCIÓN</strong>:Este fondo está cerrado</span></a></font></p>"
													'Response.Write "<p align=center><font color=gray><a class=info href='#'>" & arrAsset(I,J) & "<span><strong>PRECAUCIÓN</strong>:Este fondo está cerrado</span></a></font></p>"
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

						
						'MENCO: 09-10-2012, Beneficiarios de la Póliza de Seguro de Vida------------------------- INICIO
						'GEMINI -> 21840 - Visualización de beneficiarios y beneficio de permanencia en portal clientes y Pipeline C+S
						If TRIM(Plan) = "SKCS" Then
							Dim SqlBeneficiarios
							Dim rstBeneficiarios
							Dim arrBeneficiarios
							Dim adoBeneficiarios
							
							Set rstBeneficiarios = Server.CreateObject("ADODB.Recordset")
							Set adoBeneficiarios = GetConnpipelineDB
							
							SqlBeneficiarios = "GET_BENEFICIARIOS_POLIZA " & Contract & ", '" & TRIM(Product) & "', '" & TRIM(Plan) & "'"
							rstBeneficiarios.Open SqlBeneficiarios, adoBeneficiarios
						
							If rstBeneficiarios.BOF And rstBeneficiarios.EOF Then
								arrBeneficiarios = 0
							Else
								arrBeneficiarios = rstBeneficiarios.GetRows()
							End If
							
							rstBeneficiarios.Close
							write_sp_log adoBeneficiarios, 500, "GET_BENEFICIARIOS_POLIZA", Contract, Product, Plan, ClientId, 0, "", Session("sp_miLogin") & " at contract_info_mfund.asp"							
						
							OpenTable "100%", ""
								OpenTr ""
									OpenTd "thead", "height=15"										
									CloseTd
								CloseTr
								OpenTr ""
									OpenTd "thead", "height=30"
										Response.Write "Beneficiarios de la Póliza de Seguro de Vida Individual asociado al Programa"
									CloseTd
								CloseTr
							CloseTable
							
							OpenTable "80%", "'' border=1"
							If IsArray(arrBeneficiarios) Then
								OpenTr "class=teven"
									OpenTd "thead", " align=center"
										Response.Write "Nombre del beneficiario"
									CloseTd
									OpenTd "thead", " align=center"
										Response.Write "Parentesco"
									CloseTd
									OpenTd "thead", " align=center"
										Response.Write "Porcentaje de beneficio (%)"
									CloseTd
								CloseTr
								For I = 0 To UBound(arrBeneficiarios, 2)
									If (I Mod 2 = 0) Then
										OpenTr "class=todd"
									Else
										OpenTr "class=teven"
									End If
											OpenTd "tbody", " align=center"
												Response.Write arrBeneficiarios(0,I)
											CloseTd
											OpenTd "tbody", " align=center"
												Response.Write arrBeneficiarios(1,I)
											CloseTd
											OpenTd "tbody", " align=center"
												'MENCO: Por solicitud de J.P Maya 11-03-2013
												Response.Write FormatNumber(arrBeneficiarios(2,I),0)
											CloseTd
										CloseTr
								Next
								OpenTr "class=teven"
									OpenTd "thead", " align=center colspan='2'"
										Response.Write "Total"
									CloseTd
									OpenTd "thead", " align=center"
										Response.Write "100"
									CloseTd
								CloseTr
							End If
							CloseTable							
						Else
						
						End If
						'MENCO: 09-10-2012, Beneficiarios de la Póliza de Seguro de Vida------------------------- FIN
						
					Else
					'---ajuste contratos SIPEN at 20200512
					dim BeneficiariesSipen, objSkMtrustSipen
					Set objSkMtrustSipen = CreateObject("MSXML2.XMLHTTP")
					BeneficiariesSipen = Application("Beneficiaries") & "?contract=" & Contract
					''BeneficiariesSipen = objSkMtrustSipen.urlReader(BeneficiariesSipen)

					objSkMtrustSipen.Open "GET", BeneficiariesSipen , false
					objSkMtrustSipen.Send(null)
					BeneficiariesSipen = objSkMtrustSipen.responseText
					'---ajuste contratos SIPEN at 20200512
					'RTrim(arrContract(0,0)) <> "MFUND" ¿?¿?
						'Set objSkMtrust = Server.CreateObject("Conexion.conectar")
						'Beneficiaries = Application("Beneficiaries") & "?contract=" & Contract
						'Beneficiaries = objSkMtrust.urlReader(Beneficiaries)

						'Beneficiaries Table -- START
						OpenTable "100%", ""
							OpenTr ""
								OpenTd "", ""
									Response.Write "&nbsp;"
								CloseTd
								CloseTr
							OpenTr ""
								OpenTd "thead", ""
									Response.Write "Beneficiarios registrados"
								CloseTd
							CloseTr
						CloseTable
						OpenTable "", "'' border=1"
							Response.Write BeneficiariesSipen
						CloseTable
						'Beneficiaries Table -- END
					End If
				
					If PermiteRetiros = "N" Then
						'Retiros Table --- START
						OpenTable "50%", "'' border=0"
							OpenTr "class="
								OpenTd "'thead'", ""
									Response.Write "&nbsp;"
								CloseTd
							CloseTr
							OpenTr "class="
								OpenTd "'thead'", ""
									Response.Write "Retiros"
								CloseTd
							CloseTr					
						CloseTable
						OpenTable "90%","border=1"
							OpenTr "class=teven"
								OpenTd "thead","align=center"
									Response.Write "Este contrato NO esta habilitado para Retiros"
								CloseTd
							CloseTr
						CloseTable
					End If
					
			CloseTd
			CloseTr
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