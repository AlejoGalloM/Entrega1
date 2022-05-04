<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		process_selection.asp  102
'Path:				search/
'Created By:		A. Orozco 2001/07/23
'Last Modified:		S Soriano 2007/08/27 Comments to parameters for Retiros PAC
'			S Soriano 2007/05/30 get omni source data
'			J moreno 2004/10/19 add option hpf
'			Juan M Moreno 2003/18/09 Add Option sun
'			J Carreño 2003/09/11 Add product Alternativo
'			F. Calvache 2003/04/10  add document type
'			G Pinerez 2002/05/30
'			R. Lagos 2002/02/04
'                       A. Orozco 2001/09/04
'			A. Orozco 2001/10/08
'			Guillermo Aristizabal  2001/09/18 auth & log
'			Guillermo Aristizabal 2001/10/11
'			A. Orozco 2001/10/29
'			APC 2001/11/07 ClientId validation
'			A. Orozco 2001/11/07 MTCOR page redirection
'			A. Orozco 2001/12/07 MTCOR page redirection
'			Juan M Moreno 2003/18/09 Add Option sun
'			BArbelaez GREEN
'			Jaime A. Páez 2006/20/10 Lines para redireccionar Client/Contract Alert
'                       Get hidden field Agent Name for redirect a Client Alert
'                       Get hidden field Name State Contract for redirect a Client Alert
'			BArbelaez 20071110 Botón Disponible Discriminado,
'			equivalente a la opción Solicitud Retiro Parcial del PORTAL
'                       Armando Arias - 2008/May/09 - Cumplimiento circular SF052 - write_sp_log 13315
'Parameters:		Contract No.
'			Client Id
'			Client's name
'			Client's lastname
'           Daniel Rodriguez 2016-05-06 Se realiza validación disponibilidad de servicio OMNI GEMINI 100782
'Returns:		Results form the search page
'Additional Information:	
'===================================================================================
Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1





%>
<!--#include file="../_pipeline_scripts/url_check.asp"-->
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../_pipeline_scripts/mfundCorporativoScripts.asp"-->

<%
OpenHTML
OpenBody "loginbody", "bgColor=#ffffff leftMargin=0 topMargin=0"

Dim Num
Dim Contract, Product, Plan, ClientId, DocType, Phone, OU, Pas, PasName,idObjetivo,agentName,stateNamecontract
Dim Operation, I, objConn
Dim datosSource, arrDatosSource, SourceID, SourceName, CorpProduct, PermiteRetirosOmni, PermiteRetiros, MontoMinRetiro, SaldoMinRetiro 'Variables con informacion omni de contratos corporativos
Dim partID, planID
Dim strSql
Dim objRst
Dim arrContratoPAC,arrDatosEmpresa
Dim nombreEmpresa 'S.Soriano 2007/11/15 - Para mostrar el nombre de la empresa de los contratos PAC
Authorize 1,1
Set objConn = GetConnPipelineDB

If Request.Form("selection") <> "" Then
	Num = Request.Form("selection")
Else
	Num = Request.Form("Number")
End If

Contract = Request.Form("Contract_" & Num)
Product = Request.Form("Product_" & Num)
Plan = Request.Form("Plan_" & Num)
Pas = Request.Form("Pas_" & Num)
PasName = Request.Form("Pas_Name_" & Num)
OU = Request.Form("OU_" & Num)
ClientId = Request.Form("ClientId_" & Num)
DocType  = Request.Form("DocType_" & Num)
Phone = Request.Form("Phone_" & Num)
idObjetivo = Request.Form("idObjetivo_" & Num)
Session.Contents("ClientCity") = Trim(Request.Form("City_" & Num))
agentName = Request.Form("AgentName_" & Num)
stateNamecontract = Request.Form("NameEstCto_" & Num)

'=================================================================
' I&T - WTG 20080227 Comunicacion con RetiroSK
'=================================================================
Session.Contents("DocType") = DocType
Session.Contents("ClientId") = ClientId
Session.Contents("Contract") = Contract
Session.Contents("Product") = Product
Session.Contents("Plan") = Plan

'=================================================================
' End I&T - WTG 20080227 Comunicacion con RetiroSK
'=================================================================

'===================Obtener solo datos de la source del contrato equivalente - S.Soriano 2007/05/30============
if  VerificarPlanProducto(Application("CorpPlanes"),trim(Plan)) = true then
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
	end if


	'==Inicio Traer de la tabla temporal de contratosRetirosPAC el flag correspondiente al retiro habilitado, si existe en la tabla====
	if PermiteRetirosOmni = "N" then
		strSql = "ContratoRetirosPAC_GetForContrato " & Contract
		Set objRst = Server.CreateObject("ADODB.Recordset")
		objRst.Open strSql, objConn
		If objRst.BOF And objRst.EOF Then
			PermiteRetiros = PermiteRetirosOmni
		Else
			arrContratoPAC = objRst.GetRows()
			if arrContratoPAC(11,0) = true then
				PermiteRetiros = "Y"
			else
				if arrContratoPAC(11,0) = false then
					PermiteRetiros = "N"
				else
					PermiteRetiros = PermiteRetirosOmni
				end if
			end if
		end if
		objRst.Close
		
	
	else
		PermiteRetiros = PermiteRetirosOmni
	end if

	'====Fin Traer datos tabla temporal====
	
	'====Inicio Traer Nombre Empresa para contrato PAC=====	
	strSql = "ContratosPAC_InfoEmpresa " & Contract
	
	Set objRst = nothing
	Set objRst = Server.CreateObject("ADODB.Recordset")
	objRst.Open strSql, objConn
	If objRst.BOF And objRst.EOF Then
		nombreEmpresa = "NA"
	else
		arrDatosEmpresa = objRst.GetRows()
		nombreEmpresa = arrDatosEmpresa(5,0)
	end if
	objRst.Close
	'=====Fin Traer Nombre Empresa ========

	Session.Contents("PermiteRetiros") = PermiteRetiros
	Session.Contents("MontoMinRetiro") = MontoMinRetiro
	Session.Contents("SaldoMinRetiro") = SaldoMinRetiro
	Session.Contents("NombreEmpresa") = nombreEmpresa
else 
	datosSource = ""
end if
'===================Fin Obtener solo datos de la source del contrato equivalente - S.Soriano 2007/05/30=========

Operation = Request.Form("operation")
If Contract = "" Then Contract = 0
If (isnull(ClientId) or ClientId = "") Then ClientId = 0

write_sp_log objConn, 13315, "", Contract, Product, Plan, ClientId, 0, "", "process_selection Loaded by: " & Session("sp_miLogin") & " selected " & Operation & " for the Contract " & Contract

Select Case Operation

    Case "alertCto" ' Add JPaez
            OpenForm "menu", "post", Application("URLEditor") & "?desde=6&Cto=" & + Contract , ""
	Case "alertUser"
			OpenForm "menu", "post", Application("URLEditor") & "?desde=7", ""
	Case "rad"
			OpenForm "menu", "post", "../radication/radication.asp", ""
	Case "info"
			OpenForm "menu", "post", Application("URLEditor") & "?desde=4", ""
	Case "addProd"
			OpenForm "menu", "post", Application("URLEditor") & "?desde=5", ""
	Case "EditInv"
			OpenForm "menu", "post", Application("URLEditor") & "?desde=1", ""
	Case "EditInvNew" 
			OpenForm "menu", "post", Application("URLEditorNew") & "?desde=1", ""
	Case "MarkTBen"
			OpenForm "menu", "post", Application("URLEditor") & "?desde=13", ""
    Case "EditFatca" 
			OpenForm "menu", "post", Application("URLFATCA"), ""
	Case "CrearCE" 'GREEN
			OpenForm "menu", "post", Application("CrearCE"), ""
	'<AFILIACIÓN CLIENTE EXISTENTE>
	Case "CrearCA"
			OpenForm "menu", "post", Application("CrearCA"), ""
	'</AFILIACIÓN CLIENTE EXISTENTE>
	'PPIO - 20071110 Botón Disponible Discriminado
	Case "ConsultarDR"
			OpenForm "menu", "post", Application("ConsultarDR"), ""
	'FIN - 20071110 Botón Disponible Discriminado	
	Case "AutHPF"
			OpenForm "menu", "post", Application("URLHPF") & "?desde=0", ""
	Case "ConHPF"
			OpenForm "menu", "post", Application("URLHPF") & "?desde=1", ""
       
	Case "statement" ' Adicionado por Rlagos and JMMA
            'OpenForm "menu", "post", Application("URLStatement") & "?client=" & + ClientId + "_" + DocType , ""
			 OpenForm "menu", "post", Application("URLStatement") & "?TypeProcess=Extracto"&"&client=" & + ClientId + "_" + DocType , ""
	Case "certificado" ' Adicionado por J Ospino
            'OpenForm "menu", "post", Application("URLCertificado") & "?client=" & + ClientId + "_" + DocType , ""
		     OpenForm "menu", "post", Application("URLCertificado") & "?TypeProcess=Certificado"&"&client=" & + ClientId + "_" + DocType , ""
    '====================
	'<I&T - WTG 20081024  Key Campañas>
	'====================
	Case "CampaignsEmail"
		OpenForm "menu", "post", Application("UrlCampaigns") & "?IdCampaigns=1", ""
	'====================
	'<I&T - WTG 20081024>
	'====================
	Case Else
		Select Case Trim(Product)
			Case "SIPEN"
				OpenForm "menu", "post", "../contract_info/contract_info_mfund.asp", ""
				'OpenForm "menu", "post", Application("MVCContractInfoMfund"), ""
			Case "MFUND"
				OpenForm "menu", "post", "../contract_info/contract_info_mfund.asp", ""
				'OpenForm "menu", "post", Application("MVCContractInfoMfund"), ""
			Case "FPOB"
				OpenForm "menu", "post", "../contract_info/contract_info_fpob.asp", ""
			Case "MTCOR", "MTIND"
				OpenForm "menu", "post", "../contract_info/contract_info_mtcor.asp", ""
			Case "FPAL"
				OpenForm "menu", "post", "../contract_info/contract_info_fpal.asp", ""
			'======2003/10/30 J moreno	
			Case "FCO", "TLIFE", "IGOLD", "FONVIDA", "FIBAC", "FMAGNO", "SKINST", "OMBRAV", "OMINMA", "OMACCI", "OMLIQ", "ICGREN","ICINME","ICINMR","ICLIME","ICOPOR","ICRFGO","ICRFLO","ICRVGO","ICT108","ICT187","OMACCP","ICCATI","ICCATII","ICCAT3", "ICCATIV", "ICCATV","ICCATVI","ICDINA","IC3001","IC3652","IC3651","IC9001","IC6003"
				OpenForm "menu", "post", "../contract_info/contract_info_proto.asp", ""
			'=====Fin Modificación		
			Case "MFCR"
				OpenForm "menu", "post", "../contract_info/contract_info_mfcr.asp", ""
			Case "OMSVI", "OMPEV" ' Se agrega OMSVI al proceso - Fabian Montoya 2015/02/25  -- Se adiciona OMPEV - Fabian Montoya 2015/11/23
				OpenForm "menu", "post", "../contract_info/contract_info_seguros.asp", ""
			Case Else
				OpenForm "menu", "post", "../contract_info/contract_info_other.asp", ""
		End Select
		
	Session("contrato")=Contract

	End Select
	PlaceInput "Contract", "hidden", Contract, ""
	PlaceInput "Unit", "hidden", Request.Form("Unit_" & Num), ""
	PlaceInput "ClientId", "hidden", ClientId, ""
	
	PlaceInput "DocType", "hidden", DocType, ""
	PlaceInput "Name", "hidden", Request.Form("Name_" & Num), ""
	PlaceInput "Phone", "hidden", Phone, ""
	PlaceInput "Product", "hidden", Product, ""
	PlaceInput "Plan", "hidden", Plan, ""
	PlaceInput "Pas", "hidden", Pas, ""
	PlaceInput "PasName", "hidden", PasName, ""
	PlaceInput "AgentId", "hidden", Session("sp_IdAgte"), ""
	PlaceInput "AgentName", "hidden", agentName, ""   'Jaime
	PlaceInput "StateNameContract", "hidden", stateNamecontract, ""   'Jaime
	PlaceInput "PartID", "hidden", partID, ""
	PlaceInput "SourceID", "hidden", SourceID, ""
	PlaceInput "SourceName", "hidden", SourceName, ""
	PlaceInput "CorpProduct", "hidden", CorpProduct, ""
	PlaceInput "PermiteRetiros", "hidden", PermiteRetiros, ""
	PlaceInput "MontoMinRetiro", "hidden", MontoMinRetiro, ""
	PlaceInput "SaldoMinRetiro", "hidden", SaldoMinRetiro, ""
	PlaceInput "NombreEmpresa", "hidden", nombreEmpresa, ""

CloseForm


'Reload Left Menu -- START
OpenForm "menu_left", "post", "../menu/menu.asp", "target=menu"
	PlaceInput "Name", "hidden", Request.Form("Name_" & Num), ""
	PlaceInput "ClientId", "hidden", ClientId, ""
	PlaceInput "DocType", "hidden", DocType, ""
	PlaceInput "Phone", "hidden", Phone, ""
	PlaceInput "Contract", "hidden", Contract, ""
	PlaceInput "Product", "hidden", Product, ""
	PlaceInput "Plan", "hidden", Plan, ""
	PlaceInput "Pas", "hidden", Pas, ""
	PlaceInput "PasName", "hidden", PasName, ""
	PlaceInput "OU", "hidden", OU, ""
	
	'======  Inicio S.Soriano pendiente Paso a produccion retiros PAC ==============
	
	PlaceInput "PermiteRetiros", "hidden", PermiteRetiros, ""
	PlaceInput "MontoMinRetiro", "hidden", MontoMinRetiro, ""
	PlaceInput "SaldoMinRetiro", "hidden", SaldoMinRetiro, ""	
		
	'======== Fin S.Soriano pendiente Paso a produccion retiros PAC  ===============
CloseForm
'Reload Left Menu -- END

Response.Write "<script language=javascript>" & vbCrLf & "	document.menu.submit();" & vbCrLf
If Operation <> "info" and Operation <> "addProd" and Operation <> "EditInv" and Operation <> "rad" and operation <> "EditInvNew" and operation <> "CrearCE" and operation <> "ConsultarDR" Then
	Response.Write "	document.menu_left.submit();" & vbCrLf
End If
Response.Write "</SCRIPT>" & vbCrLf
CloseBody
CloseHTML

If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>