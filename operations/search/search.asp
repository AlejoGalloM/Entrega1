<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		search.asp 100
'Path:				search/
'Created By:		Andres Felipe Orozco July 17, 2001
'Last Modified:		Fabio Calvache
'			Jul 30 marathon Add document type
'			A. Orozco 2001/09/21
'			A. Orozco 2001/10/08
'	                Html layout, Write to log
'			Guillermo Aristizabal  2001/09/18 auth & log
'			Guillermo Aristizabal 2001/10/11
'			A. Orozco 2001/12/27
'			First Name Field is no longer mandatory
'			Juan M moreno 2003/18/09 add access to SUN
'			Javier vargas search operativo  of operating and not operating (yaxa eu) 21/jun/2005
'                       Armando Arias - 2008/May/09 - Cumplimiento circular SF052 - PlaceTitle - write_sp_log 13316
'Parameters:		Contract No. or
'			Client Id or
'			Client's name or
'			Client's lastname
'			Diana Mariced Pérez - 2009/02/04 - Modificación requerida por el proceso de Referencia Unica de Recaudos
'			Ampliación longitud del #contrato de 7 a 12.
'Returns:
'Additional Information:Contained inside the frameset ../main/main.htm
'===================================================================================
Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../_pipeline_scripts/wsbls_scripts.asp"-->
<!--#include file="../_pipeline_scripts/wsbls_scripts.asp"-->
<!--#include file="../_pipeline_scripts/url_check.asp"-->
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<%

	'		If Session.Contents("SiteRetuns") = "2" Then
	'			PlaceAnchor "cerrar.asp""" & _
	'			" class=lmenu onClick=""JavaScript:return confirm('Confirma que desea salir?')"" target=""_parent", "Salir"
	'		Else
	'			PlaceAnchor "../../login/login.asp?Out=1""" & _
	'			" class=lmenu onClick=""return confirm('Confirma que desea salir?') "" target=""_parent", "Salir"
	'		End If
	'		'</I&T-WTG 20080115>

Authorize 1,1
Dim objConn
Dim objRst
Dim strSQL
Dim Combo, combo2
Dim ComboDocType
Dim ComboDocTypeFP
Dim Contract, ClientId, Name, LastName
Dim Init_msg
dim fsoImage,Path,Img
dim text
Set objConn = GetConnPipelineDB
dim wsbls
Dim Cuenta
'<I&T - WTG 20070921 >
Dim UtilitarianString

Session("name")=""
Session("contrato")=""

UtilitarianString = " OnChange = JavaScript:DocumentType();"
'</I&T>
'Build product drop down
Combo = "<SELECT name=Product class=listaLogin onChange='javascript:document.searchContract.btnContract.focus()'>" & vbCrLf & _
"<OPTION value=''>-- Producto --</OPTION>" & vbCrLf
strSQL = "spsp_GetComboProducts"

Set objRst = Server.CreateObject("ADODB.Recordset")
objRst.Open strSQL, objConn
Do Until objRst.EOF
	Combo = Combo & _
	"<OPTION value='" & objRst.Fields("Producto") & "'>" & _
	objRst.Fields("Descripcion") & "</OPTION>" & vbCrLf
	objRst.MoveNext
Loop
objRst.Close

Combo2 = "<SELECT name=Product class=listaLogin onChange='javascript:document.searchContract.btnContract.focus()'>" & vbCrLf & _
"<OPTION value=''>-- Producto --</OPTION>" & vbCrLf
strSQL = "spsp_GetComboProducts"
Set objRst = Server.CreateObject("ADODB.Recordset")
objRst.Open strSQL, objConn
Do Until objRst.EOF
	Combo2 = Combo2 & _
	"<OPTION value='" & objRst.Fields("Producto") & "'>" & _
	objRst.Fields("Descripcion") & "</OPTION>" & vbCrLf
	objRst.MoveNext
Loop
objRst.Close

'Build Document type drop down Fabio Calvache Mayo 2003
ComboDocType = "<SELECT name=DocType " & UtilitarianString & " class=listaLogin>" & vbCrLf & _
			   "<OPTION value='O'>-- Tipo de Documento --</OPTION>" & vbCrLf

ComboDocTypeFP  = "<SELECT name=DocTypeFp " & UtilitarianString & " class=listaLogin>" & vbCrLf & _
				  "<OPTION value='O'>-- Tipo de Documento --</OPTION>" & vbCrLf

strSQL = "DocumentType_GetAllDesc"
Set objRst = Server.CreateObject("ADODB.Recordset")
objRst.Open strSQL, objConn
Do Until objRst.EOF
	ComboDocType = ComboDocType & "<OPTION value='" & objRst.Fields("Tipo") & "'>" & objRst.Fields("Descripcion") & "</OPTION>" & vbCrLf
	ComboDocTypeFP = ComboDocTypeFP & "<OPTION value='" & objRst.Fields("Tipo") & "'>" & objRst.Fields("Descripcion") & "</OPTION>" & vbCrLf	
	objRst.MoveNext
Loop
ComboDocType = ComboDocType & "</SELECT>"
ComboDocTypeFP = ComboDocTypeFP & "</SELECT>"
objRst.Close
write_sp_log objConn, 13316, "spsp_GetComboProducts", 0, "", "", 0, 0, "", "search.asp Loaded by " & Session("sp_miLogin")
Combo = Combo & "</SELECT>"

Session.Contents("sp_aboutpl")   = False
%>
<html>
	<head>
		<meta name="" http-equiv="expires" content="Wednesday, 27-Dec-95 05:29:10 GMT"/>
		<meta name="" http-equiv="Pragma" content="no_cache"/>
		<title>Buscar</title>
		<script languaje="javascript" src='../_pipeline_scripts/validation.js'></script>
		<script language="javascript">
			function DocumentType()
			{
				var typeId = document.searchClientId.DocType.value;
				document.searchClientId.hdTypeDocument.value = typeId;
			}
		</script>
		<link href="../../css/OLDMutualStyle.css" rel="stylesheet" type="text/css"/>
</head>
<body onLoad='document.searchContract.txtContrato.focus()' class="cuerpo">
    <!--div class="encabezado">			//GV Se procede a eliminar por proyecto marca.
            Pipeline
    </div>
    <div class="rounded">
        <b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b>
    </div-->
	</br></br></br></br></br></br></br>
	<div class="contenido">
		<!--div class="subtituloPagina">      //GV Se procede a eliminar por proyecto marca.
			Busqueda Cliente - Contrato
		</div-->
		<table align="center" class="tblBusquedaPrincipal">
			<tr>
				<td>
					<table class="tblBusquedaPrincipal" align="center">
					<%
						If Session("pwdchg") = 1 Then
							OpenTr ""
								OpenTd "texto-informativo", "colspan=4"
									Response.Write "<br>" & vbCrLf & _
									"Su Password ha sido cambiado satisfactoriamente.<br><br>" & vbCrLf
									Session("pwdchg") = 0
								CloseTd
							CloseTr
						End If
						If Request.Form("NoResult") = 1 Then
							write_sp_log objConn, 13316, "", 0, "", "", 0, 0, "", "search.asp Loaded by " & Session("sp_miLogin") & " no results for the search."
							OpenTr ""
								OpenTd "texto-informativo", "colspan=4"
									If Request.Form("txtFpId") <> "" Then
										Response.Write "<br>" & vbCrLf & _
										"No se encontraron resultados para el FP identificado " & ucase(Request.Form("txtFpId")) & "<BR><BR> Por favor refine su búsqueda."
									else
										Response.Write "<br>" & vbCrLf & _
										"No se encontraron resultados para la persona <h5>" & ucase(Request.Form("Name"))  & " " & ucase(Request.Form("LastName"))  & "</h5> Por favor refine su búsqueda"
										If autorizarMn(1,6) Then
											Response.Write " o adicione el cliente o el empleador." & vbCrLf
										Else
											Response.Write "." & vbCrLf
										End If
									End If
									Response.Write "<br><br>" & vbCrLf
								CloseTd
							CloseTr
							If autorizarMn(4,25) Then
								If Request.Form("txtFpId") = "" Then		
									OpenForm "searchEditor", "post", Application("UrlEditor")+ "?desde=0", "onSubmit='formValidation(this)'" 
										OpenTr ""
											OpenTd "texto-informativo", "colspan=4"
											PlaceInput "btnEditar", "submit", "   Adicionar Cliente    ", " class='button-OLD btnBusqueda'"
											PlaceInput "ClientID", "hidden", Request.Form("ClientId"), "id='CI'"
											PlaceInput "DocType", "hidden", Request.Form("DocType"), "id='Type'" 
										CloseTd
										CloseTr
									closeform
								End If
							End if 
							If autorizarMn(1,6) Then
								If Request.Form("txtFpId") = "" Then		
									OpenForm "searchEmpleador", "post", Application("UrlEditor") + "?desde=2", "onSubmit='formValidation(this)'" 
									OpenTr ""
										OpenTd "thead", "colspan=4"
											PlaceInput "btnEditarempleador", "submit", "   Adicionar Empleador    ", " class='button-OLD btnBusqueda'"
											PlaceInput "ClientID", "hidden", Request.Form("ClientId"), "id='CI'"
											PlaceInput "DocType", "hidden", Request.Form("DocType"), "id='Type'" 
										CloseTd
									CloseTr
									closeform
								End If
							End If
						End If
						CloseConnPipelineDB
						OpenTr ""
							OpenTd "subtitulo", "'' align=left colspan=4"
								Response.Write "<br>" & vbCrLf & "<p>Consulte aquí el  número de contrato o Referencia de pago</p>" & vbCrLf
							CloseTd	
						CloseTr
						OpenTr ""
							OpenTd "separadorSecciones", "colspan=4"
								Response.Write "&nbsp;"
							CloseTd	
						CloseTr
						OpenForm "searchReference", "post", "search_results.asp?", "onSubmit='return formValidation(this)'"
							OpenTr "align=left"
								OpenTd "tblBusquedaColum0", "align=left"
									PlaceInput "btnReference", "submit", "Obtener Referencia/Contrato", "id='               Enviar' class='button-OLD btnBusqueda'"
								CloseTd
								OpenTd "tbodyBusqueda", "nowrap align=right"
									PlaceInput "txtContratos", "text", Session.Contents("ContratoRU"), "maxlength=10 id='XW       I     Contrato No.' class=txbox"
									PlaceInput "sp", "hidden", "spem_GetReferenciaUnica", "id='               SP'"
									Session.Contents("ContratoRU") = ""
									Response.Write Combo2
								CloseTd
							CloseTr
						    if session("Reference")<>"" then
								OpenTr ""
									OpenTd "thead", "align=center"
									CloseTd
									OpenTd "tbody", "nowrap align=right"
										PlaceInput "txtReferencia", "text", Session.Contents("Reference"), "maxlength=12 id='XW       I     Contrato No.' class=txbox"
										Session.Contents("Reference") = ""
										CloseTd
									OpenTd "", ""
									CloseTd					  		
								CloseTr
							end if
						CloseForm
						if (Session.Contents("SiteRetuns") = "2")then
                        OpenTr ""
								OpenTd "texto-informativo", "colspan=4"
									Response.Write "<br>" & vbCrLf & _
									"Por favor consultar el cliente desde Portal Distribuidores....<br><br>"
                                    response.End 
								CloseTd
							CloseTr
						end if
						OpenTr ""
							OpenTd "subtitulo", "colspan=4"
								Response.Write "<br>" & vbCrLf		
								Response.Write "Búsquedas" & vbCrLf
							CloseTd
						CloseTr
						OpenTr ""
							OpenTd "separadorSecciones", "colspan=4"
								Response.Write "&nbsp;"
							CloseTd					  
						CloseTr
						OpenForm "searchContract", "post", "search_results.asp?", "onSubmit='return formValidation(this)'"
							OpenTr ""
								OpenTd "thead", "align=left"
									PlaceInput "btnContract", "submit", "    Buscar x Contrato    ", "id='               Enviar' class='button-OLD btnBusqueda'"
								CloseTd
								OpenTd "tbodyBusqueda", "nowrap align=right"
									PlaceInput "txtContrato", "text", "", "maxlength=12 id='XW       I     Contrato No.' class=txbox"
                                    Response.Write "  "
									PlaceInput "sp", "hidden", "spsp_SearchContract", "id='               SP'"
									Response.Write Combo
								CloseTd
							CloseTr
						CloseForm
						OpenTr ""
							OpenTd "", "colspan=4"
								Response.Write "&nbsp;"
							CloseTd					  
						CloseTr
						OpenForm "searchClientId", "post", "search_results.asp?", "onSubmit='return formValidation(this)'"
							OpenTr ""
								OpenTd "thead", "align=left"
									PlaceInput "btnClientId", "submit", "Buscar x Identificación", "id='               Enviar' class='button-OLD btnBusqueda'"
								CloseTd
								OpenTd "tbodyBusqueda", "nowrap align=right"
									PlaceInput "hdTypeDocument", "hidden", "", ""
									PlaceInput "txtIdCliente", "text", "", "id='RA       I     Número Identificación' class=txbox"
									PlaceInput "sp", "hidden", "spsp_SearchClientId", "id='               SP'"
									Response.Write ComboDocType
								CloseTd					  		
							CloseTr
						CloseForm
						OpenTr ""
							OpenTd "", "colspan=4"
								Response.Write "&nbsp;"
							CloseTd					  
						CloseTr	
						OpenForm "searchNames", "post", "search_results.asp?", "onSubmit='return formValidation(this)'"
							OpenTr ""
								OpenTd "thead", "align=left"
									PlaceInput "btnClientName", "submit", "   Buscar x Nombres/Apellidos   ", "id='               Enviar' class='button-OLD btnBusqueda'"
								CloseTd
								OpenTd "tbodyBusqueda", "nowrap align=right"
                                %>
                                    <div>
                                        <input name="txtNombres" class="txbox tbxnombres" id="             L  Nombres" type="text"/>
                                        /
                                        <input name="txtApellidos" class="txbox tbxnombres" id="R            L  Apellidos" type="text"/>
                                        <input name="sp" id="               SP" type="hidden" value="spsp_SearchNames"/>
                                    </div>
                                <%
								CloseTd
							CloseTr
						CloseForm
						OpenTr ""
							OpenTd "", "colspan=4"
								Response.Write "&nbsp;"
							CloseTd					  
						CloseTr
						OpenForm "searchFpId", "post", "search_results.asp?", "onSubmit='return formValidation(this)'"
							OpenTr ""
								OpenTd "thead", "align=left"
									PlaceInput "btnFpId", "submit", "   Buscar x Id Fp   ", "id='               Enviar' class='button-OLD btnBusqueda'"
								CloseTd
								OpenTd "tbodyBusqueda", "nowrap align=right"
									PlaceInput "txtFpId", "text", "", "id='RA       I     Número Identificación FP' class=txbox"
                                    Response.Write "  "
									PlaceInput "sp", "hidden", "spsp_SearchFpId", "id='               SP'"
									Response.Write ComboDocTypeFP
								CloseTd
							CloseTr
						CloseForm
						OpenTr ""
							OpenTd "", "colspan=4"
								Response.Write "&nbsp;"
							CloseTd					  
						CloseTr
						If autorizarMn(5,27) Then
							OpenTable "''", "width=100%  border=0 align=center height=15% "
								OpenTR ""
									OpenTd "thead", "width=100% align=left"
										Response.Write "&nbsp;"
									CloseTD		
								CloseTr
							closetable
							response.write session("str")
							closetable
							closetd
							closetr
							closetable
						end if
			'Reload Left Menu -- START
			OpenForm "menu", "post", "../menu/menu.asp", "target=menu"
				PlaceInput "Name", "hidden", "", ""
				PlaceInput "ClientId", "hidden", "", ""
				PlaceInput "Contract", "hidden", "", ""
				PlaceInput "Product", "hidden", "", ""
				PlaceInput "Plan", "hidden", "", ""
				PlaceInput "Option", "hidden", 2, ""
			CloseForm
%>
		</div>
		<script language="javascript">
			document.menu.submit();
		</script>
		<!--div class="rounded">			//GV Se procede a eliminar por proyecto marca.
			<b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b>
		</div-->		
	</body>
</html>
<%
'Reload Left Menu -- END
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
