<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		search_results.asp  101
'Path:				search/
'Created By:		A. Orozco 2001/07/17
'Last Modified:		S Soriano 2007/10/25 add PartID(cuenta individual) if corporate contract
'			S Soriano 2007/05/29 add source name if corporate contract
'			Juan Manuel Moreno 2003/10/19	add options to HPF autorizar estudio vincular contrato
'			Fabio Calvache April 10 / 2033 Add Document type
'			Juan M Moreno	2003/09/18 added buttons for interact with SUN and modify Selection contract
'			R. Lagos 2003/09/23 change document type in N
'			GAR - GAP 2002/05/23 Add Pas Name
'			R. Lagos 2002/02/05 Add Pas
'                       A. Orozco 2001/09/21
'			A. Orozco 2001/10/08
'			Guillermo Aristizabal  2001/09/18 auth & log
'			Guillermo Aristizabal 2001/10/11
'			A. Orozco 2001/10/29
'			A. Orozco 2001/12/10 added log registration for search stored procedure
'			Jaime A. Páez 2006/20/10 Add button for redirect a Client Alert and Contract Alert
'                       Add hidden field to Agent Name for redirect a Contract Alert
' ****DEBE ACTIVARSE****Add hidden field Name State Contract for redirect a Client Alert
'			BArbelaez GREEN
'                       Armando Arias - 2008/May/09 - Cumplimiento circular SF052 - PlaceTitle - write_sp_log 13317
'			I&T - WTG 2009/01/06 Inclusión de búsqueda por lista vinculante por número de contrato
'			Camilo Gutierrez I&T 11-02-2011 Agregar Columnas Risk Profile( Perfil del Inversionista ) & Real Contract Profile ( Portafolio de Inversion )
'Parameters:		Contract No. or
'			Client Id or
'			Client's name or
'			Client's lastname
'			Diana Mariced Pérez Corzo - 2009/02/05 - Modificación Referencia Unica - Invocación de sp spem_GetReferenciaUnicaInversa>
'			Phanor Torres I&T 03-03-2015 Cambiar nombre a Risk Profile( Perfil del Inversionista ) por Perfil de Inversión del contrato / inversionista, 
'                                        y se modifica el proceso de visualizar el perfil, incluyendo los productos FICS
'			Phanor Torres I&T 20-05-2015 Cambiar tooltip a RiskProfile para contratos con productos FICS o CONCOM
'Returns:		Results form the search page
'Additional Information:Contained inside the frameset ../main/main.htm
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
<!--#include file="../_pipeline_scripts/WSBLS_Scripts.asp"-->
<!--#include file="../../operations/transactions/reglamentoscriptsPipeline.asp"-->

<SCRIPT LANGUAGE="javascript">

<!--
    function ShowFP(agent, product) {
		window.open('fp_details.asp?agent=' + agent + '&product=' + product, '', 'toolbar=no,scrollbars=no,width=320,height=200')
	}

	function ShowHelp(page) {
		window.open(page, '', 'toolbar=no,scrollbars=no,width=650,height=900')
	}
	
//-->

</SCRIPT>
<%
Authorize 1,1
dim Page
dim placeAnchor2
Dim wsbls
Dim xmlDOC
Dim strSql 'Stored procedures and SQL queries
Dim objConn 'ADODB Connection
Dim objRst 'ADODB Recordset
Dim IdCliente, DocType
Dim IdFp, DocTypeFp
Dim arrContracts, arrUnits , arrContratoRef'Array - search results used to store recordsets data
Dim Contrato, Prod, Pas, PasName
Dim I, J, K, L 'Used to navigate arrays
Dim QString 'Used to send the form data as a querystring when there are no results
Dim Total 'Display total of records
Dim Name, LastName, Socs, arrSocs, Flag, TotalRadioBtns
Dim planCorp, NombreEmpresa 'Para guardar plan si es corporativo
Dim datosSource, arrDatosSource 'Datos de la source omni equivalente al contrato ulla
Dim arrDatosEmpresa 'S.Soriano, para mostrar nombre de la empresa
dim ContratoRef
dim arrReference  ' <I&T - DMPC>
dim mensaje  ' <I&T - DMPC>
dim Producto ' <I&T - DMPC>
dim Reference  ' <I&T - DMPC>
dim alfinn
dim hasContractInfo
hasContractInfo = false
dim hasContractFICSOrCONCOM ' Variable added to the definition of investor profile on 2015/03/03 by Phanor Torres
hasContractFICSOrCONCOM =false 
dim getCaptureRiskProfile
getCaptureRiskProfile= false ' Variable added to the definition of risk profile on 2015/05/20 by Phanor Torres
Dim productsFICSOrCONCOM ' Define the number of products to work, added to deploy risk profile by Phanor Torres 2015/03/03
Dim itemFics
dim hasContractMarkingTaxBenefit ' Variable que define si se muestra la opcion de beneficio tributario on 2015/06/30 Fabian Montoya
hasContractMarkingTaxBenefit =false 
Dim productsMarkingTaxBenefit ' Carga los productos a los que aplica la marcacion de beneficio tributario 2015/06/30 Fabian Montoya
Dim itemMarking
Dim sqlLista
Dim arrClientContract
Dim ViewOFAC
Dim strContract
dim NewEditorAut
Dim userAffected, spaceAffected

spaceAffected=""
NewEditorAut = false
Set objConn = GetConnpipelineDB
Select Case Request.Form("SP")
	Case "spsp_SearchContract"
		If len(Request.Form("txtContrato"))<12 then
			If Request.Form("txtContrato") = "" Then
				Contrato = 0
			Else
				Contrato = Request.Form("txtContrato")
			End If
			If Request.Form("Product") <> "" Then
				strSql = "spsp_SearchContract " & Contrato & ", '" & Request.Form("Product") & "'"
			Else
				strSql = "spsp_SearchContract " & Contrato
			End If
		spaceAffected=Replace(strSql,"'","")
		Else
			'==========================================================================
			''<I&T - DMPC 2009/02/11 - Modificado por proceso de Referencia Unica>
			' Valida el número de contrato digitado, revisando el dígito de verificación
			'==========================================================================
			If Request.Form("txtContrato") = "" Then
				ContratoRef = 0
			Else
				ContratoRef = Request.Form("txtContrato")
			End If
			If Request.Form("Product") <> "" Then
				Producto =Request.Form("Product")
			Else
				Producto="null"
			End If
			arrReference = GetReferenciaUnicaInversa (ContratoRef, Producto)
			mensaje=arrReference(2,0)
			If Mensaje<>"null" Then
				Response.Write("<SCRIPT LANGUAGE='javascript'>") 
				Response.Write("alert('" + mensaje +"');") 
				Response.Write("window.location.href = 'search.asp';")
				Response.Write("</SCRIPT>") 
			Else
				Contrato=arrReference(0,0)
				Producto=arrReference(1,0)
				strSql = "spsp_SearchContract " & Contrato & ", '" & Producto & "'"
			End If 
			spaceAffected=Replace(strSql,"'","-")
		End if
	Case "spem_GetReferenciaUnica"
		'==========================================================================
		''<I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica - Recaudos>
		' Obtiene la referencia única a partir del Contrato y el producto
		'==========================================================================
		spaceAffected=spaceAffected&" - "&"spem_GetReferenciaUnica"
		If Request.Form("txtContratos") = "" Then
			Contrato = 0
		Else
			Request.Form("txtContratos")
			If IsNumeric(Request.Form("txtContratos")) then
				contrato= Request.Form("txtContratos")
				session("ContratoRU")=Request.Form("txtContratos")
			Else
				Response.Write("<SCRIPT LANGUAGE='javascript'>") 
				Response.Write("alert('El número de Contrato no puede contener letras, ni signos. Intente nuevamente.');") 
				Response.Write("window.location.href = 'search.asp';")
				Response.Write("</SCRIPT>") 
			End if
		End If
		If Request.Form("Product") <> "" Then
			Producto =Request.Form("Product")
		Else
			Producto="null"
			Response.Write("<SCRIPT LANGUAGE='javascript'>") 
			Response.Write("alert('Para obtener la Referencia Única debe digitar el número de contrato y seleccionar un producto del listado. Intente nuevamente.');") 
			Response.Write("window.location.href = 'search.asp';")
			Response.Write("</SCRIPT>") 
		End If
		Reference = GetReferenciaUnica( Producto, Contrato)
	   	If Reference <> "null" Then
			session("Reference")=Reference
			response.Redirect("search.asp")
	   	Else
			Response.Write("<SCRIPT LANGUAGE='javascript'>") 
			Response.Write("alert('El producto seleccionado no está configurado para Referencia Única.');") 
			Response.Write("window.location.href = 'search.asp';")
			Response.Write("</SCRIPT>")
		End If 
	Case "spsp_SearchClientId"
		Contrato = 0
		If Request.Form("txtIdCliente") = "" Then
			IdCliente = 0
			DocType = "O"
		Else
			IdCliente = Request.Form("txtIdCliente")
			DocType = Request.Form("DocType")
		End If
		userAffected=IdCliente
		strSql = "spsp_SearchClientIdAndType " & ReadNumber(IdCliente) & ", '" & DocType & "'"
		spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")		
	Case "spsp_SearchNames"	
		Contrato = 0
		Name = Request.Form("txtNombres")
		LastName = Request.Form("txtApellidos")
		strSql = "spsp_SearchNames '" & Name & "', '" & LastName & "'" 
		spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
		userAffected=Name&" "&LastName
	Case "spsp_SearchFpId"  'Adicionado por Alejandro Jaramillo / Pasantias Julio 2012
		Contrato = 0
		IdCliente = 0
		If Request.Form("txtFpId") = "" Then
			IdFp = 0
			DocTypeFp = "O"
		Else
			IdFp = Request.Form("txtFpId")
			DocTypeFp = Request.Form("DocTypeFp")
		End If
		strSql = "spsp_SearchFpIdAndType " & ReadNumber(IdFp) & ", '" & DocTypeFp & "'"
		spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
		userAffected= IdFp
	End Select

	Select Case Session("sp_miLogin")
		Case "gap" 'Skandia
				NewEditorAut = true		
	End Select
	
	If Contrato = "" Then Contrato = 0
	If IdCliente = "" Then IdCliente = 0
	write_sp_log objConn, 13317, Replace(strSQL, "'", "''"), Contrato, "", "", ReadNumber(IdCliente), 0, "", "SP_LOG - Start Search " & Session("sp_miLogin")
	Set objRst = Server.CreateObject("ADODB.Recordset")
	objRst.Open strSql, objConn
	If objRst.BOF And objRst.EOF Then
		arrContracts = 0
		write_sp_log objConn, 13317, Replace(strSQL, "'", "''"), Contrato, "", "", ReadNumber(IdCliente), 0, "", "SP_LOG - No results " & Session("sp_miLogin")
	Else
		arrContracts = objRst.GetRows()
		write_sp_log objConn, 13317, Replace(strSQL, "'", "''"), Contrato, "", "", ReadNumber(IdCliente), 0, "", "SP_LOG - Results Found: " & objRst.RecordCount & " " & Session("sp_miLogin")
	End If
	objRst.Close
	'===============================================================================
	'Check SDN table
	strSql = "spsp_GetSDNPeople "
	If IdCliente <> "" And IdCliente <> "0" Then
		strSql = strSql & "'', '', " & ReadNUmber(IdCliente)
	Else
		strSql = strSql & "'" & Name & "', '" & LastName & "'"
	End If
	spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
	If Contrato = 0 Then
		session("Sp_Name")= Name
		session("Sp_LastName")= LastName
		session("IdCliente")= ReadNUmber(IdCliente)
		strSQL = Replace(strSQL, "', '", "-")
		strSQL = Replace(strSQL, "'", "")
		'Validar si la consulta se hizo por alguno de estos campos
		If IdCliente <> "" Or Name <> "" Or LastName <> "" Then
			xmldoc = GetNombres_Busqueda(Session("sp_Name") + " " + Session("sp_LastName"), Session("IdCliente") )
			'Validar que no haya devuelto error el WS
			If xmldoc <> "ERROR TRAYENDO DATOS DEL CLIENTE" Then
				Dim ofac		
				ofac = split(xmldoc,",")
				i=0
				If (Len(xmldoc) >0) Then
					Response.Write "<SCRIPT LANGUAGE=javascript>" & vbCrLf & _
					"<!--" & vbCrLf & _
					"window.parent.showModalDialog('add_client/sdn_list.asp','sdn','channelmode=no, titlebar=no, toolbar=no, status=no, menubar=no, scrollbars=yes, width=550, heigth=500, top=0, left=100');" & vbCrLf & _
					"//-->" & vbCrLf & _
					"</SCRIPT>" & vbCrLf
				End If
			End If
		End If
	Else
		If IsArray(arrContracts) Then
			sqlLista = "spsp_SearchClientByContract " & Contrato
			If Request.Form("Product") <> "" Then
				sqlLista = sqlLista & ",'" & Request.Form("Product") & "'"
			End If
			spaceAffected=spaceAffected&" - "&Replace(sqlLista,"'","-")
			objRst.Open sqlLista, objConn
			If objRst.BOF And objRst.EOF Then
				arrClientContract = 0
			Else
				arrClientContract = objRst.GetRows()
			End If		
			objRst.Close
			ViewOFAC = False
			If IsArray(arrClientContract) Then
				For I = 0 To UBound(arrClientContract, 2)
					xmldoc = GetNombres_Busqueda( arrClientContract(4,I), CStr(arrClientContract(0,I)))
					If xmldoc <> "" And xmldoc <> "ERROR TRAYENDO DATOS DEL CLIENTE" Then
						ViewOFAC = True
					End If
				Next
				If ViewOFAC Then
					strContract = "?Contract=" & Contrato & "&Product=" & Request.Form("Product")
					Response.Write "<SCRIPT LANGUAGE=javascript>" & vbCrLf & _
					"<!--" & vbCrLf & _
					"window.parent.showModalDialog('add_client/sdn_list.asp" & strContract & "','sdn','channelmode=no, titlebar=no, toolbar=no, status=no, menubar=no, scrollbars=yes, width=550, heigth=500, top=0, left=100');" & vbCrLf & _
					"//-->" & vbCrLf & _
					"</SCRIPT>" & vbCrLf
				End If
			End If
		End If
	End If
	strSql = Replace(strSql, "', '", "-")
	strSql = Replace(strSql, "'", "")
	write_sp_log objConn, 13317, strSql, Contrato, "", "", ReadNumber(IdCliente), 0, "", "add_client.asp no matches were found in the SDN list by " & Session("sp_miLogin")
	'Check No Elegibles table
	strSql = "Spsp_GetSDNPeople_Noelegibles "
	If IdCliente <> "" And IdCliente <> "0" Then
		strSql = strSql & "'', '', " & ReadNumber(IdCliente)
	Else
		strSql = strSql & "'" & Name & "', '" & LastName & "'"
	End If
	spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
	If Contrato = 0 Then
		If (IdCliente <> "" And IdCliente <> "0") Or Name <> "" Or LastName <> "" Then
			Set objRst = Server.CreateObject("ADODB.Recordset")
			objRst.Open strSql, objConn
			If Not(objRst.BOF And objRst.EOF) Then
				Session("sp_sqlSDN") = strSQL
				strSQL = Replace(strSQL, "', '", "-")
				strSQL = Replace(strSQL, "'", "")
				write_sp_log objConn, 13317, strSQL, Contrato, "", "", ReadNumber(IdCliente), 0, "", "add_client: A match was found in the SDN list by " & Session("sp_miLogin")
				VentanaModal "Cliente No Elegible","Favor debe advertir inmediatamente al oficial de cumplimiento sobre la información de este cliente a ingresar."
			End If
			objRst.Close
		End If
	End If
	strSql = Replace(strSql, "', '", "-")
	strSql = Replace(strSql, "'", "")
	write_sp_log objConn, 13317, strSql, Contrato, "", "",  ReadNumber(IdCliente), 0, "", "add_client.asp no matches were found in the SDN list by " & Session("sp_miLogin")
	'===============================================================================
	'Get whole saler's societies
	If Session("sp_AccessLevel") = "1" Then
		strSql = "spsp_GetWSSocieties '" & Session("sp_Idworker") & "'"
		objRst.Open strSql, objConn
		If objRst.BOF And objRst.EOF Then
			Socs = 0
		Else
			arrSocs = objRst.GetRows()
		End If
		objRst.Close
	End If
	spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
	strSql = Replace(strSql, "', '", "-")
	strSql = Replace(strSql, "'", " ")
	write_sp_log objConn, 13317, strSql, Contrato, Prod, "", 0, 0, "", "search_results.asp Loaded by " & Session("sp_miLogin")
%>
<html>
	<head>
        <title>Resultados Buscar</title>
        <meta name="" http-equiv="expires" content="Wednesday, 27-Dec-95 05:29:10 GMT"/>
        <meta name="" http-equiv="Pragma" content="no_cache"/>
        <link href="../../css/OLDMutualStyle.css" rel="stylesheet" type="text/css"/>
        <script type="text/javascript" src="../_pipeline_scripts/jquery-1.10.2.min.js">
        </script>	
        <script>
            $(document).ready(function () {
                $('#AceptMessage').click(function () {
                    $('.Popup').fadeOut(200);
                });
            });
        </script>
        <script type="text/javascript" language="javascript">
		
			function setSession(buttonGroup){
				var docType = getSelectedDocumentType(buttonGroup);
				var docNumber = getSelectedDocumentNumber(buttonGroup);
				var contractId = getSelectedContract(buttonGroup);
				var product = getSelectedProductType(buttonGroup);
				var userName = document.cookie.replace(/(?:(?:^|.*;\s*)sp%5Flogin\s*\=\s*([^;]*).*$)|^.*$/, "$1");
								
				
				const sessionData = {
				"DocumentType": docType,
				"DocumentNumber": docNumber,
				"ContractId": contractId,
				"Product": product,
				"ModifiedBy": userName
			  }
			  sessionStorage.setItem("sessionData", window.btoa(unescape(encodeURIComponent(JSON.stringify(sessionData)))));
			}
		
            function getSelectedRadio(buttonGroup) 
            {
                // returns the array number of the selected radio button or -1 if no button is selected
                if (buttonGroup[0]) 
	            { // if the button group is an array (one button is not an array)
                    for (var i=0; i<buttonGroup.length; i++) 
                    {
                        if (buttonGroup[i].checked) 
                        {
			                return i
                        }
                    }
	            } 
                else 
	            {
	                if (buttonGroup.checked) 
                    {
                        return 0; 
                    } // if the one button is checked, return zero
	            }
                // if we get to this point, no radio button is selected
	            return -1;
            } // Ends the "getSelectedRadio" function
    
            function getSelectedRadioValue(buttonGroup) 
            {
                // returns the value of the selected radio button or "" if no button is selected
                var i = getSelectedRadio(buttonGroup);
                if (i == -1) 
                {
                    return "";
                }
                else 
                {
                    if (buttonGroup[i]) 
                    { 
                        // Make sure the button group is an array (not just one button)
                        return buttonGroup[i].value;
                    }
                    else 
                    { 
                        // The button group is just the one button, and it is checked
                        return buttonGroup.value;
                    }
                }
            } // Ends the "getSelectedRadioValue" function
			
			function getSelectedDocumentType(buttonGroup) 
            {
			
		        var collection;
		        var j = -1;
				
                var n = getSelectedRadio(buttonGroup);
				
		        radioName = "DocType_"+n;
		        			
                // returns the value of the selected radio button or "" if no button is selected
                if (n == -1) 
                {
                    return "";
                }
                else 
                {
                    collection = document.transaction.elements;
					for (i = 0; i < collection.length; i++) 
					{
						if (collection[i].name == radioName)
						{
							return(collection[i].value);
						}	
					}
                }
				return "";
            } // Ends the "getSelectedDocumentType" function

			function getSelectedDocumentNumber(buttonGroup) 
            {
			
		        var collection;
		        var j = -1;
				
                var n = getSelectedRadio(buttonGroup);
				
		        radioName = "ClientId_"+n;
		        			
                // returns the value of the selected radio button or "" if no button is selected
                if (n == -1) 
                {
                    return "";
                }
                else 
                {
                    collection = document.transaction.elements;
					for (i = 0; i < collection.length; i++) 
					{
						if (collection[i].name == radioName)
						{
							return(collection[i].value);
						}	
					}
                }
				return "";
            } // Ends the "getSelectedDocumentType" function
			
			function getSelectedProductType(buttonGroup) 
            {
			
		        var collection;
		        var j = -1;
				
                var n = getSelectedRadio(buttonGroup);
				
		        radioName = "Product_"+n;
		        			
                // returns the value of the selected radio button or "" if no button is selected
                if (n == -1) 
                {
                    return "";
                }
                else 
                {
                    collection = document.transaction.elements;
					for (i = 0; i < collection.length; i++) 
					{
						if (collection[i].name == radioName)
						{
							return(collection[i].value.trim());
						}	
					}
                }
				return "";
            } // Ends the "getSelectedProductType" function
			
			function getSelectedContract(buttonGroup) 
            {
			
		        var collection;
		        var j = -1;
				
                var n = getSelectedRadio(buttonGroup);
				
		        radioName = "Contract_"+n;
		        			
                // returns the value of the selected radio button or "" if no button is selected
                if (n == -1) 
                {
                    return "";
                }
                else 
                {
                    collection = document.transaction.elements;
					for (i = 0; i < collection.length; i++) 
					{
						if (collection[i].name == radioName)
						{
							return(collection[i].value);
						}	
					}
                }
				return "";
            } // Ends the "getSelectedDocumentType" function
			
            function getRadioValue() 
            {
		        var collection;
		        var j = -1;
		        alert("inicio");	
		        radioName = "Number";
		        collection = document.transaction.elements;
		        for (i = 0; i < collection.length; i++) 
                {
                    if (collection[i].type == "radio" && collection[i].name == radioName)
                    {
			            j = j + 1;
			            if (collection[i].checked)
    				        return(j);
	    	        }	
		        }
		        return j;
	        }
        </script>
        <script type="text/javascript" language="javascript" src='../_pipeline_scripts/validation.js'></script>
	</head>
    <body class="cuerpo">
        <div class="encabezado">
            Pipeline
        </div>
        <div class="rounded">
            <b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b>
        </div>
	    <div class="contenido">
		    <div class="subtituloPagina">
			    Contrato - Cliente
		    </div>
<%
If IsArray(arrContracts) Then 'Build Table With Data
%>
    <form name="transaction" onsubmit="return formValidation(this)" action="process_selection.asp" method="post">
        <input name="operation" type="hidden"/>
        <input name="selection" type="hidden"/>
<%
	Total = UBound(arrContracts, 2) + 1
    %>
    <div class="tblContenido">
	<table class="tblValores tblContenido" border="0">
        <thead>
    <%
        OpenTr ""
            OpenTd "'titulotabla'", "colspan=12"
                response.Write "Contratos Clientes"
            closetd
        closetr
        OpenTr ""
            OpenTd "'separadorSecciones'", "colspan=12"
            closetd
        closetr		
        OpenTr ""
			OpenTd "'texto-informativo'", "colspan=12"
				Response.Write "Total: " & Total & " registros"
			CloseTd
		CloseTr
		OpenTr "class=thead align=center"
			If autorizarMn(1,6) Then
			OpenTh "", ""
				Response.Write "Radicar"
			End If
			OpenTh "", ""
				Response.Write "Sel"
			CloseTh
			OpenTh "", ""
				Response.Write "Producto/Plan"
			CloseTh
			OpenTh "", ""
				Response.Write "Contrato"
			CloseTh
			'=============================================================================================
			'Added By Camilo Gutierrez I&T 11-02-2011 Display Risk Profile & Real Contract Profile Columns
			'Modified By Phanor Torres I&T 03-03-2015 Display Risk Profile & Real Contract Profile Columns
			'=============================================================================================
			OpenTh "", ""
				Response.Write "Perfil de Inversión del contrato / inversionista"
			CloseTh
			OpenTh "", ""
				Response.Write "Portafolio del inversionista"
			CloseTh
			OpenTh "", ""
				Response.Write "Identificación"
			CloseTh
			OpenTh "", ""
				Response.Write "Tipo"
			CloseTh
			OpenTh "", ""
				Response.Write "Cliente"
			CloseTh
			OpenTh "", ""
				Response.Write "UO"
			CloseTh
'=============================================================================================
'Added By A. Orozco 2001/10/25
'Display FP and Society
'=============================================================================================
			OpenTh "", ""
				Response.Write "Agente Comercial"
			CloseTh
'=============================================================================================
'Added By J. Páez 2007/01/25
'Display Compliance Alert 
'=============================================================================================
			If autorizarMn(3,40) Then '  
				OpenTh "", ""
					Response.Write "Alerta"
				CloseTh
			end if
'=============================================================================================
'=============================================================================================
		CloseTr
        %>
        </thead>
        <%
	TotalRadioBtns = 0
	Pas = "N" 
	PasName = ""
	
	hasContractMarkingTaxBenefit =false
	For J = 0 To UBound(arrContracts, 2)
		If (J Mod 2) = 0 Then
			OpenTr "class=filaSombra align=center"
		Else
			OpenTr "class=filaBlanca align=center"
		End If
		'Alejandro Figueroa - Proyecto Advice Tools - 2011-03-16
		'La funcionalidad aplica para todos los contratos Multifund Individual y para los Corporativo tipo CAHC (los demás no)
		Dim Product
		Dim Plan
		Product = arrContracts(5,J)
		Plan = arrContracts(6,J)

		hasContractInfo = false		
		if (trim(Product) = "MFUND" or trim(Product) = "CONCOM") and VerificarPlanProducto(Application("CorpPlanesAll"),trim(Plan)) <> True Then
			hasContractInfo = true
		End if
		'=============================================================================================
		'Added By Phanor Torres I&T 03-03-2015
		'Display Risk Profile
		'=============================================================================================
		hasContractFICSOrCONCOM = false
		getCaptureRiskProfile = false
		productsFICSOrCONCOM = split(Application("ProductsFICSOrCONCOM"),";")
		
		For Each itemFics In productsFICSOrCONCOM 
			if(trim(Product) = itemFics) then
				hasContractFICSOrCONCOM = true
			End if						
		Next
		
		if(hasContractInfo <> true) then
			if(hasContractFICSOrCONCOM = true) then
				getCaptureRiskProfile = true
			End if						
		Else
			getCaptureRiskProfile = true
		End if
		'=============================================================================================
		
		'=============================================================================================
		'Added By Fabian Montoya I&T 30-06-2015
		'Display Marking Tax Benefit
		'=============================================================================================
		productsMarkingTaxBenefit  = split(Application("productsMarkingTaxBenefit"),";")	
		
		For Each itemMarking In productsMarkingTaxBenefit 
			if(trim(Product) = itemMarking) then
				hasContractMarkingTaxBenefit = true
			End if						
		Next
		
		'=============================================================================================
		
		'======================================================================================
		'Start modify by R. Lagos add Validation Pas
		'Feb. 1/2002
		'last  modified by GAR- GAP 
		'2002/05/23
		'======================================================================================
			If  Trim(arrContracts(5,J)) = "MFUND" OR Trim(arrContracts(5,J)) = "MTCOR" Then
				strSql = "exec spsp_getPasName '" & Trim(arrContracts(5,J)) & "'," & Trim(arrContracts(7,J))
				objRst.Open strSql, objConn
				If not objRst.BOF and not objRst.EOF Then
					Pas = "S"
					PasName = objRst(0)
				else
					Pas = "N"   
					PasName = ""			   
				end if
					objRst.Close
				spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
			End If 
		'=======================================================================================		
			If autorizarMn(1,6) Then
			OpenTd "", ""
				PlaceInput "rad", "image", "", "src='../../images/operations/boton.gif' id='" & J & "'" & _
				" onclick='javascript:transaction.operation.value=""rad""; transaction.selection.value=" & J & "'"
			CloseTd
			End If
			OpenTd "", ""
				PlaceInput "Contract_" & J, "hidden", arrContracts(7,J), ""
				PlaceInput "Unit_" & J, "hidden", arrContracts(10,J), ""
				PlaceInput "ClientId_" & J, "hidden", arrContracts(2,J), ""
				PlaceInput "DocType_"  & J, "hidden", arrContracts(17,J), ""
				PlaceInput "Product_" & J, "hidden", arrContracts(5,J), ""
				PlaceInput "Plan_" & J, "hidden", arrContracts(6,J), ""
				PlaceInput "OU_" & J, "hidden", arrContracts(10,J), ""
				'============================================
				' I6T - WTG 20080221 Visualizacón correcta 
				'	    del nombre de una empresa por concatenación
				'============================================
				If arrContracts(17,J) <> "N" Then
					Name = arrContracts(1,J) & " " & arrContracts(0,J)
				Else
					Name = LTrim(arrContracts(1,J)) & LTrim(arrContracts(0,J))
				End If
				'============================================
				' END I6T - WTG 20080221 
				'============================================
				PlaceInput "Name_" & J, "hidden", Name, ""
				PlaceInput "Phone_" & J, "hidden", arrContracts(12,J), ""
				PlaceInput "City_" & J, "hidden", arrContracts(16,J), ""
				PlaceInput "AgentId_" & J, "hidden", Session("sp_IdAgte"), ""
				PlaceInput "AgentName_" & J, "hidden", arrContracts(13,J), "" 'Jaime
				PlaceInput "NameEstCto_" & J, "hidden", arrContracts(3,J), "" 'Jaime
				PlaceInput "Pas_" & J, "hidden", Pas, ""
				PlaceInput "Pas_Name_" & J, "hidden", PasName, ""
				
				If IsNull(arrContracts(5,J)) Or arrContracts(5,J) = "" Then
					If J = 0 Then
								PlaceInput "Number", "radio", J, "checked id='          R   '"
							Else
								PlaceInput "Number", "radio", J, " id='          R   '"
							End If
					TotalRadioBtns = TotalRadioBtns + 1
				Else
				Select Case Session("sp_AccessLevel")
					Case 0 'Skandia
							If J = 0 Then
								PlaceInput "Number", "radio", J, "checked id='          R   '"
							Else
								PlaceInput "Number", "radio", J, " id='          R   '"
							End If
							TotalRadioBtns = TotalRadioBtns + 1
					Case 1 'WHOLE SALER					
						Flag = False
						For K = 0 To UBound(arrSocs,2)
							If Not(IsNull(arrContracts(9,J))) Then
								If CStr(arrSocs(0,K)) = CStr(arrContracts(9,J)) Then
									Flag = True
									Exit For
								Else
									Flag = False
								End If
							End If
						Next
						If Flag Then
							If J = 0 Then
								PlaceInput "Number", "radio", J, "checked id='          R   '"
							Else
								PlaceInput "Number", "radio", J, " id='          R   '"
							End If
							TotalRadioBtns = TotalRadioBtns + 1
						Else
							Response.Write "N/A"
						End If
			'=============================================================================================
			'Modificado Por Julian Zapata 2016-02-01
			'Permitir seleccionar para consultar contratos de las sociedades que hacen parte de la estuctura de una sociedad Intermediaria
			'=============================================================================================
					Case 2 'Partner FP, Fran. Worker
						If Not(IsNull(arrContracts(9,J))) Then
							If CStr(Session("sp_idSoc")) = CStr(arrContracts(9,J)) AND Session("esAgenteIntermediario") = False Then
								If J = 0 Then
									PlaceInput "Number", "radio", J, "checked id='          R   '"
								Else
									PlaceInput "Number", "radio", J, " id='          R   '"							
								End If
								TotalRadioBtns = TotalRadioBtns + 1
							Else
								 If ((CStr(Session("sp_idSoc")) = CStr(arrContracts(9,J)) OR CStr(Session("sp_idSoc")) = CStr(arrContracts(20,J)) OR CStr(Session("sp_idSoc")) = CStr(arrContracts(21,J)) ) AND Session("esAgenteIntermediario") = True) Then
									If J = 0 Then
										PlaceInput "Number", "radio", J, "checked id='          R   '"
									Else
										PlaceInput "Number", "radio", J, " id='          R   '"							
									End If
									TotalRadioBtns = TotalRadioBtns + 1
								 Else
									Response.Write "N/A"
								 End If
							 End If
						Else
							Response.Write "N/A"
						End If
					Case 3 'FP
						If CStr(Session("sp_IdAgte")) = CStr(arrContracts(8,J)) Then
							If J = 0 Then
								PlaceInput "Number", "radio", J, "checked id='          R   '"
							Else
								PlaceInput "Number", "radio", J, " id='          R   '"
							End If
							TotalRadioBtns = TotalRadioBtns + 1
						Else
							Response.Write "N/A"
						End If
					Case 4 'Anyone else
						Response.Write "N/A"
				End Select
				End If
			CloseTd
'=============================================================================================
'Added By Fabio Calvache 2003/04/09
'Display Product
'=============================================================================================
			OpenTd "", ""
				If IsNull(arrContracts(5,J)) Or arrContracts(5,J) = "" Then
					Response.Write "N/A"
				Else
					planCorp = arrContracts(6,J)					
					dim auxArrDatosSource, auxDatosSource
					Dim objRec
					Dim Descciption
					Dim IsSkandiaGreen
					Set objRec = Server.CreateObject("ADODB.Recordset")
					IsSkandiaGreen = false
					'Se consulta si es MFUND y si es Skandia Green
					If  RTrim(arrContracts(5,J)) = "MFUND" Then
						'Busca si el contrato es Green
						strSql = "spem_GetDescriptionFondosCerrados " & arrContracts(7,J) &",'" & RTrim(arrContracts(5,J)) &"','" & RTrim(arrContracts(6,J)) &"'"
						objRec.Open strSql, objConn
						If not objRec.BOF and not objRec.EOF Then
							Descciption = objRec(0)
						Else
							Descciption = ""			   
						End if
						objRec.Close
						spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
						'Busca si el contrato es Green
						If ucase(RTrim(Descciption)) = "MFUND" Or Descciption = "" or ucase(RTrim(Descciption)) = "MULTIFUND" Then
							IsSkandiaGreen = False
						Else
							Descciption = UCase(Descciption)
							IsSkandiaGreen = True
						End If
					End If
					If  VerificarPlanProducto(Application("CorpPlanes"),trim(planCorp)) = True Then
						auxDatosSource = GetSourceInfo(arrContracts(7,J))	
						'====S.Soriano Inicio Traer Nombre Empresa para contrato PAC=====	
						strSql = "ContratosPAC_InfoEmpresa " & arrContracts(7,J)
						Set objRec = nothing
						Set objRec = Server.CreateObject("ADODB.Recordset")
						objRec.Open strSql, objConn
					If objRec.BOF And objRec.EOF Then
						nombreEmpresa =  "NA"
						else
						arrDatosEmpresa = objRec.GetRows()
						nombreEmpresa = arrDatosEmpresa(5,0)
						end if
						objRec.Close
						spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
						'=====S.Soriano Fin Traer Nombre Empresa ========						
						auxArrDatosSource = split(auxDatosSource,"-")
						If InStr(1,auxArrDatosSource(0),"ERROR") = 0  Then
							Response.Write  arrContracts(5,J) & "<br>" & arrContracts(6,J) & "<br>" & auxArrDatosSource(4) & "<br>" & nombreEmpresa
						Else
							If  IsSkandiaGreen Then
								Response.Write  Descciption
							Else
								Response.Write  arrContracts(5,J) & "<br>" & arrContracts(6,J) & "<br>"
							End If
						End If
					Else
						If  IsSkandiaGreen Then
							Response.Write  Descciption 
						Else
							Response.Write  arrContracts(5,J) & "<br>" & arrContracts(6,J)
						End If
					End If
				End If
			CloseTd
			OpenTd "", ""
'==========================================================================
''<I&T - DMPC 2009/02/13 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia única a partir del Contrato y el producto
'==========================================================================
				dim Referencia
				dim ContratoFCO
				If Trim(arrContracts(5,J)) = "FCO" Then
					ContratoFCO = Trim(CStr(arrContracts(6,J)))+ Trim(CStr(arrContracts(7,J)))
					Reference = GetReferenciaUnica( arrContracts(5,J), ContratoFCO)
				else
					Reference = GetReferenciaUnica( arrContracts(5,J), arrContracts(7,J))
				end if
				If Not Isnull(arrContracts(6,J)) Then
					If VerificarPlanProducto(Application("CorpPlanPatrocinado"), Trim(CStr(arrContracts(6,J)))) = True then
						Reference = "PATROCINADO [" & Reference & "]"
					End if
				End If
				If Not(IsNull(Reference)) Then
					Referencia= Reference
				else
					Referencia= arrContracts(7,J)
				end if
'==========================================================================		
				If IsNull(arrContracts(7,J)) Or arrContracts(7,J) = "" Then
					Response.Write "N/A"
				Else
					if VerificarPlanProducto(Application("CorpPlanes"),trim(planCorp)) = True  Then 'S.Soriano 2007/10/25
						auxArrDatosSource = split(auxDatosSource,"-")
						
						if InStr(1,auxArrDatosSource(0),"ERROR") = 0 then
							Response.Write  "<b>" & Referencia & "</b>" & "<br>" & "<b>" & auxArrDatosSource(0) & "<br>" & auxArrDatosSource(3) &  "</b>" & "<br>" & arrContracts(3,J) 
						else
						Response.Write  "<b>" & Referencia & "</b>" & "<br>" & arrContracts(3,J)
						end if
					else
						'======================================================
						'Camilo Gutierrez I&T 10-02-2011 Agregar Nombre de contrato
						'======================================================
						If  (RTrim(arrContracts(5,J)) = "MFUND" OR RTrim(arrContracts(5,J)) = "CONCOM") Then
							strSql = "sppl_GetDetallesContrato " & arrContracts(7,J) & ", '" & arrContracts(5,J) & "', '" & arrContracts(6,J) & "'"
							Dim rstContractInfo
							Dim arrContractInfo
							Set rstContractInfo = Server.CreateObject("ADODB.Recordset")
							rstContractInfo.Open strSql, objConn
							If rstContractInfo.BOF And rstContractInfo.EOF Then
								arrContractInfo = 0
								Response.Write "<b>" & Referencia & "</b>" & "<br>" & arrContracts(3,J)
							Else
								arrContractInfo = rstContractInfo.GetRows()
								Dim strNombreContrato
								strNombreContrato = ""
								If hasContractInfo and not isnull(arrContractInfo(103,0)) and arrContractInfo(103,0) <> "" Then
									strNombreContrato = arrContractInfo(103,0) & "<br />"
								End If
								Response.Write  strNombreContrato & "<b>" & Referencia & "</b>" & "<br/>" & arrContracts(3,J)
							End If
							rstContractInfo.Close
							spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
						Else
							Response.Write "<b>" & Referencia & "</b>" & "<br>" & arrContracts(3,J)
						End If
						'======================================================
						'Camilo Gutierrez I&T 10-02-2011 Agregar Nombre de contrato
						'======================================================						
					end if			
				End If
			CloseTd
			'=============================================================================================
			'=============================================================================================
			'Added By Camilo Gutierrez I&T 11-02-2011
			'Modified By Phanor Torres
			'Display Risk Profile I&T 03-03-2015
			'=============================================================================================
			OpenTd "", ""
				If getCaptureRiskProfile  Then
					strSql = "Relacionamiento..Contract_Profile_GetProfileByContractByCountry '" & trim(arrContracts(17,J)) & trim(arrContracts(2,J)) & "', 'CO', " & arrContracts(7,J) & ", '" & trim(arrContracts(5,J)) & "'"
					Dim rstRiskProfile
					Dim arrRiskProfile
					Dim strRiskProfile
					Set rstRiskProfile = Server.CreateObject("ADODB.Recordset")
					rstRiskProfile.Open strSql, objConn

					If rstRiskProfile.BOF And rstRiskProfile.EOF Then
						arrRiskProfile = 0

						If hasContractFICSOrCONCOM then
							strRiskProfile = "No definido. <img title='Este perfil es definido por el cliente en el portal  o a través del formato físico de Encuesta de Perfil de Riesgo y Categorización de inversionista o asignado de acuerdo a las inversiones de su cliente para los productos administrados por Old Mutual Fiduciaria S.A. y Old Mutual Valores S.A. Sociedad Comisionista de Bolsa.' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
						Else
							strRiskProfile = "No definido. <img title='Este perfil es definido por su cliente en la Encuesta de Perfil de Inversión del formato de vinculación Multifund o a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
						
						End If
					Else
						arrRiskProfile = rstRiskProfile.GetRows()
						If isnull(arrRiskProfile(2,0)) or arrRiskProfile(2,0) = "" Then
							If hasContractFICSOrCONCOM then
								strRiskProfile = "No definido. <img title='Este perfil es definido por el cliente en el portal  o a través del formato físico de Encuesta de Perfil de Riesgo y Categorización de inversionista o asignado de acuerdo a las inversiones de su cliente para los productos administrados por Old Mutual Fiduciaria S.A. y Old Mutual Valores S.A. Sociedad Comisionista de Bolsa.' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
							Else
								strRiskProfile = "No definido <img title='Este perfil es definido por su cliente en la Encuesta de Perfil de Inversión del formato de vinculación Multifund o a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
							End If
						Else
							If hasContractFICSOrCONCOM then
								strRiskProfile = arrRiskProfile(2,0) & " <img title='Este perfil es definido por el cliente en el portal  o a través del formato físico de Encuesta de Perfil de Riesgo y Categorización de inversionista o asignado de acuerdo a las inversiones de su cliente para los productos administrados por Old Mutual Fiduciaria S.A. y Old Mutual Valores S.A. Sociedad Comisionista de Bolsa.' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
							Else
								strRiskProfile = arrRiskProfile(2,0) & " <img title='Este perfil es definido por su cliente en la Encuesta de Perfil de Inversión del formato de vinculación Multifund o a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
							End If
						End If
					End If
					Response.Write  strRiskProfile
					rstRiskProfile.Close
					spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
				Else
					Response.Write "N/A"
				End If
			CloseTd
			'=============================================================================================
			'Added By Camilo Gutierrez I&T 11-02-2011
			'Display Real Profile
			'=============================================================================================
			OpenTd "", ""
				If hasContractInfo Then
					strSql = "Relacionamiento..Contract_Profile_GetRealRiskProfileByContractByCountry '" & trim(arrContracts(17,J)) & trim(arrContracts(2,J)) & "', 'CO', " & arrContracts(7,J) & ", '" & trim(arrContracts(5,J)) & "'"
					Dim rstRealProfile
					Dim arrRealProfile
					Dim strRealProfile
					Set rstRealProfile = Server.CreateObject("ADODB.Recordset")
					rstRealProfile.Open strSql, objConn
					If rstRealProfile.BOF And rstRealProfile.EOF Then
						arrRealProfile = 0
						strRealProfile = "No definido <img title='Este perfil es calculado a partir de la distribución real de la inversiones de su cliente en los diferentes portafolios' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
					Else
						arrRealProfile = rstRealProfile.GetRows()
						If isnull(arrRealProfile(2,0)) or arrRealProfile(2,0) = "" Then
							strRealProfile = "No definido <img title='Este perfil es calculado a partir de la distribución real de la inversiones de su cliente en los diferentes portafolios' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
						Else
							strRealProfile = arrRealProfile(2,0) & " <img title='Este perfil es calculado a partir de la distribución real de la inversiones de su cliente en los diferentes portafolios' src='../../images/operations/InfoIcon.png' style='margin-top: -3px' />"
						End If
					End If
					Response.Write  strRealProfile
					rstRealProfile.Close
					spaceAffected=spaceAffected&" - "&Replace(strSql,"'","")
					Response.Write "N/A"
				End If
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(2,J)) Or arrContracts(2,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write arrContracts(2,J)
				End If
			CloseTd
			'=============================================================================================
			'Added By Fabio Calvache 2003/04/09
			'Display Type Id
			'=============================================================================================
			OpenTd "", ""
				If IsNull(arrContracts(17,J)) Or arrContracts(17,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write arrContracts(17,J)
				End If
			CloseTd
			'=============================================================================================
			'Added By A. Orozco 2001/10/25
			'Display FP and Society
			'=============================================================================================
			OpenTd "", ""
				If arrContracts(17,J) <> "N" Then
					If IsNull(arrContracts(0,J)) Or IsNull(arrContracts(1,J)) Or arrContracts(0,J) = "" Or  arrContracts(1,J) = "" Then
						Response.Write "N/A"
					Else
						Response.Write arrContracts(0,J) & "<br>" & arrContracts(1,J)
					End If
				else
					If IsNull(arrContracts(1,J)) Or  arrContracts(1,J) = "" Then
						Response.Write "N/A"
					Else
						'============================================
						' I6T - WTG 20080221 Visualizacón correcta 
						'	    del nombre de una empresa por concatenación
						'============================================
						Response.Write arrContracts(1,J) & arrContracts(0,J)
						'============================================
						' END I6T - WTG 20080221
						'============================================
					End If
				end if	
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(10,J)) Or arrContracts(10,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write arrContracts(10,J)
				End If
			CloseTd
			'=============================================================================================
			'Added By A. Orozco 2001/10/25
			'Display FP and Society
			'=============================================================================================
			OpenTd "", ""
				If IsNull(arrContracts(13,J)) Or arrContracts(13,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write "<a href='javascript:ShowFP(""" & arrContracts(8, J) & """,""" & trim(Product) & """);'>" & arrContracts(13,J) & "</a>"
				End If
			CloseTd
			'=============================================================================================
			'Added By J. Páez 2007/01/25
			'Display Compliance Button Contract Alert 
			'=============================================================================================
			If autorizarMn(3,40) Then 
			OpenTd "", ""
   				If arrContracts(18,J) = True Then
					PlaceInput "alertCto", "image", "", "src='../../images/operations/ContractMark.gif' id='" & J & "'" & _
					" onclick='javascript:transaction.operation.value=""alertCto""; transaction.selection.value=" & J & "'"
				else
					PlaceInput "alertCto", "image", "", "src='../../images/operations/NonAlert.gif' id='" & J & "'" & _
					" onclick='javascript:transaction.operation.value=""alertCto""; transaction.selection.value=" & J & "'"
				End If
			CloseTd
			End If
		CloseTr
	Next
	OpenTr ""
		OpenTd "''", "colspan=12 align=center"
			PlaceInput "select", "submit", "Seleccionar Contrato", "class=button-OLD"
	
	'=============================================================================================
			'Modificado Por Julian Zapata 2016-02-01
			'Ocultar Botones para Sociedades Canal Intermediario
			'=============================================================================================
	If autorizarMn(1,1) and Session("esAgenteIntermediario") = False  Then
		 PlaceInput "selectSun", "submit", "Información cliente", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='info'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 		
	end if
	If autorizarMn(5,25) and Session("esAgenteIntermediario") = False Then
		PlaceInput "selectSunProduct", "submit", "Adicionar Producto cliente", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='addProd';  var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	end if 
	'generar extracto 
	If autorizarMn(2,33) and Session("esAgenteIntermediario") = False Then
		PlaceInput "selectextracto", "submit", "Generar Extracto", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='statement'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	End If
	'generar certificado 
	If autorizarMn(2,33) and Session("esAgenteIntermediario") = False Then
		PlaceInput "selectextracto", "submit", "Generar Certificado", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='certificado'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	End If
	If autorizarMn(6,23) and Session("esAgenteIntermediario") = False Then
		PlaceInput "selectSunProduct", "submit", "Editar cliente", "class=button-OLD onclick=" & chr(34) & "javascript: transaction.operation.value = 'EditInv';var sel = getSelectedRadioValue(document.transaction.Number);if (sel == ''){alert('Seleccione un cliente');return false;}transaction.selection.value = sel;form.submit()" & chr(34) 
	End if
	If autorizarMn(6,23) and NewEditorAut and Session("esAgenteIntermediario") = False  Then
		PlaceInput "selectSunProduct", "submit", "Nuevo editor cliente", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='EditInvNew'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	End if
    If autorizarMn(6,23) and Session("esAgenteIntermediario") = False Then
		PlaceInput "selectSunProduct", "submit", "Información FATCA", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='EditFatca'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	End if
	If autorizarMn(1,29) and Session("esAgenteIntermediario") = False Then
		PlaceInput "selectHpFAut", "submit", "Autorizar HPF", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='AutHPF'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	End if
	If autorizarMn(2,29) and Session("esAgenteIntermediario") = False Then
		PlaceInput "selectHpFCon", "submit", "Vincular Estudio HPF", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='ConHPF'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	End if
	If autorizarMn(6,47) and hasContractMarkingTaxBenefit and Session("esAgenteIntermediario") = False  Then
		PlaceInput "MarkingTaxBenefit", "submit", "Marcación para beneficio tributario", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='MarkTBen'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
	End if
	'=============================================================================================
	' Added By J. Páez (I&T) Add button CLIENT alert
	'=============================================================================================
	If autorizarMn(3,40) Then 
		PlaceInput "alertUser", "submit", "ALERTA Cliente", "class=button-OLD onclick=" &chr(34) & 	"javascript:transaction.operation.value='alertUser'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} 	transaction.selection.value= sel " & chr(34) 
	End If
	'=============================================================================================
	' Added By IBM Add button Complementacion
	'=============================================================================================
	If autorizarMn(6,23) and Session("esAgenteIntermediario") = False and arrContracts(17,0)="N" Then
		PlaceInput "Complement", "button", "Complementacion", "class=button-OLD onclick=" & chr(34) & "setSession(document.transaction.Number);location.href = '../corporate/form.asp';" & chr(34) 
	End if
	'====================
	'<I&T - WTG 20081024  Key Campañas>
	'====================
	IF CBool(Application("ViewCampaigns")) Then
				If autorizarMn(6,23) Then
					PlaceInput "selectEmailCampaigns", "submit", "Campaña Email", "class=button-OLD onclick=" & chr(34) & "javascript:transaction.operation.value='CampaignsEmail'; var sel= getSelectedRadioValue(document.transaction.Number);if(sel==''){alert('Seleccione un cliente'); return false;} transaction.selection.value= sel " & chr(34) 
				End if
	End If
    CloseTd
	CloseTr

	write_dataLog Response.Status,"search_results.asp","SP_LOG - Start Search " & Session("sp_miLogin"),userAffected,spaceAffected,"N/A","null","Consulta","N/A"		

	'====================
	'<I&T - WTG 20081024>
	'====================
    %>
    </table>
    </div>
    </form>
    <%
	If TotalRadioBtns = 0 Then 'Disable Selection buttons
		Response.Write "<script language=javascript>" & vbCrLf & _
		"<!--" & vbCrLf & _
		"	sendForm(document.transaction)" & vbCrLf & _
		"//--></script>"
	End If
	'If there's only one result, go directly to contract_info
	If UBound(arrContracts, 2) = 0 And TotalRadioBtns > 0 Then
	Else 'Continue as is
		'Reload Left Menu -- START
		OpenForm "menu", "post", "../menu/menu.asp", "target=menu"
			PlaceInput "Contract", "hidden", Contrato, ""
			PlaceInput "product", "hidden", Producto, ""
			PlaceInput "Option", "hidden", 3, ""
		CloseForm
%>
<script language="javascript">
    document.menu.submit();
</script>
<%
		'Reload Left Menu -- END
	End If
Else
	CloseConnpipelineDB
	If Err.number <> 0 Then
		Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
		"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
		Server.URLEncode(Err.source))
	End If
	OpenForm "SDN", "post", "search.asp", ""
		PlaceInput "Name", "hidden", Name, ""
		PlaceInput "LastName", "hidden", LastName, ""
		PlaceInput "ClientId", "hidden", IdCliente, ""
		PlaceInput "DocType", "hidden", DocType, ""
		PlaceInput "Contrato", "hidden", Contrato, ""
		PlaceInput "Product", "hidden", Producto, ""
		PlaceInput "NoResult", "hidden", "1", ""
		PlaceInput "DocTypeFp", "hidden", DocTypeFp, ""
		PlaceInput "txtFpId", "hidden", IdFp, ""
	CloseForm
	Response.Write "<script language=javascript>" & vbCrLf & _
	"	document.SDN.submit();" & vbCrLf & _
	"</SCRIPT>" & vbCrLf
End If
%>
        </div>
        <div class="rounded">
			<b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b>
		</div>	
	</body>
</html>
<%
CloseConnpipelineDB
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>