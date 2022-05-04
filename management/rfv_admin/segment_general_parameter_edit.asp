<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../../operations/_pipeline_scripts/url_check.asp"-->
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<%
'===================================================================================
'File Name:		segment_general_parameter.asp
'Path:				management/rfv_admin
'Created By:		G. Pinerez 2004/04/14
'Last Modified:	
'						
'Modifications:	
'Parameters:		none
'
'Returns:			Mangement default page
'Additional Information:
'ToDoes : "Permitir modificar los parametros generales"
'===================================================================================
On Error Resume Next
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
Dim adoConn
Dim strSQL
Dim arrParameters
Dim objRst 'Recordset object de los parametros 
Dim objRstAccionEnPArametro 'Recordset object de los parametros 
Dim Modificando
Dim Eliminando
Dim Adicionando

Dim Clasificacion
Dim ComboParametros
Dim idParametro
Dim Parametro
Dim Descripcion
Dim Contador
Dim ClassCss
Dim idParametroModificar
Dim processInfo, component_id, valueLog

Contador = 0
Parametro = Request.QueryString("Parametro")
idParametro = Request.QueryString("idParametro")
Descripcion = Request.QueryString("Descripcion")

Modificando = Request.Form("Modificando")
Eliminando = Request.Form("Eliminando")
Adicionando = Request.Form("Adicionando")

' ***************** POR HACER ****
' Authorize 0,8			AUTORIZACION !!!!!!!!!!!
' *****************
Set adoConn = GetConnPipelineDB
' ***************** POR HACER ****
'	DETERMINAR EL LOG !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'	write_sp_log adoConn, 8700, "", 0, "", "", 0, 0, "", "management/default.asp Loaded by " & Session("sp_miLogin")
' ***************** 
'Set adoConn = Server.CreateObject("ADODB.Connection")


component_id = "segment_general_parameter_edit.asp"
processInfo =  ""

valueLog="Parametro: "&Parametro&", idParametro: "&idParametro&", Descripcion:"&Descripcion

if Adicionando = "True" then	
	idParametroModificar = Request.Form("idParametroModificar")
	adoConn.Execute "exec trfv..Segmentacion_AddParametro '" &  Request.Form("Valor") & "','" & Parametro & "'"
 	write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "exec trfv.dbo.sprfv_getParameters - trfv..Segmentacion_GetParametros'" & Trim(strType) & "'"&"exec trfv..Segmentacion_AddParametro '" &  Request.Form("Valor") ,"",valueLog,"Operación-Adicionar","N/A"
end if


if Eliminando = "True" then	
	idParametroModificar = Request.Form("idParametroModificar")
	adoConn.Execute "exec trfv..Segmentacion_DeleteParametro " & idParametroModificar 
	 	write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "trfv..Segmentacion_DeleteParametro- trfv..Segmentacion_GetParametros '" & Trim(strType) & "'"&"exec trfv..Segmentacion_AddParametro '" &  Request.Form("Valor") ,"",valueLog,"Operación-Eliminación","N/A"
end if

if Modificando = "True" then	
	idParametroModificar = Request.Form("idParametroModificar")
	adoConn.Execute "exec trfv..Segmentacion_SetParametro  " & idParametroModificar & ",'" & Request.Form("Valor") & "'"
 	write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "exec trfv..Segmentacion_SetParametro - trfv..Segmentacion_GetParametros '" & Trim(strType) & "'"&"exec trfv..Segmentacion_AddParametro '" &  Request.Form("Valor"), "",valueLog,"Operación-Modificación","N/A"
end if

Set objRst = Server.CreateObject("ADODB.Recordset")
strSQL = "exec trfv..Segmentacion_GetParametros '" & Parametro & "'"
objRst.Open strSQL, adoConn, 3

on error resume next

OpenHTML
OpenHead
PlaceMeta "Pragma", "", "no_cache"
PlaceLink "REL", "stylesheet", "../../css/OLDMutualStyle.css", "text/css"
%>
<script language="javascript" src="../../operations/_pipeline_scripts/validation.js"></script>
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
            Segmentaci�n de Clientes
		</div>
<%

otbl"tblcontenido"
    opentr""
        opentd"''",""
            otbl"tblvalores"
	            OpenTr "valign=middle"
		            OpenTd "'titulotabla'", "align=center valign=middle colspan=2"
			            Response.Write "Administraci�n de par�metros generales - " & Request.QueryString("Descripcion") 
		            CloseTd
	            CloseTr
            ctbl
        closetd
    closetr
    opentr""
        opentd"''",""
            otbl"tblvalores"
		            OpenTr ""			
			            OpenTd "'texto-informativo'", "align=center colspan=4"
				            Response.Write " Tenga precauci�n ya que algunos par�metros deben aparecer una sola vez "
			            CloseTd
			            OpenTd "''", ""
			            CloseTd
		            CloseTr
                    opentr""
                        OpenTd "''", ""
                            Response.Write "<br><br>"
		                CloseTd
                    closetr
                    Do Until objRst.EOF
		                if (Contador mod 2) <> 0 then
			                ClassCss = "filaSombra"
		                else
			                ClassCss = "filaBlanca"
		                end if
		                Contador = Contador + 1
		            OpenTr ""
		                OpenForm "consi", "post", "segment_general_parameter_edit.asp?Parametro=" _ 
		                & Parametro & "&idParametro=" & idParametro & "&Descripcion=" & Descripcion , _
		                "onSubmit='javascript:return formValidation(this)'"
			                OpenTd ClassCss, ""
				                PlaceInput "Valor", "TextBox", objRst.Fields("Valor") , " class=txboxGenericas"				
			                CloseTd
			                OpenTd ClassCss, "align=left"
				                PlaceInput "Modificar", "submit", "Modificar", "class=button-OLD"
			                CloseTd
                             PlaceInput "Modificando", "hidden", "True", ""	
                             PlaceInput "idParametroModificar", "hidden", objRst.Fields("id"), "class=button-OLD"
				             PlaceInput "Parametro", "hidden", Parametro, "class=button-OLD"
		                CloseForm
		                OpenForm "consi", "post", "segment_general_parameter_edit.asp?Parametro=" _ 
		                & Parametro & "&idParametro=" & idParametro & "&Descripcion=" & Descripcion , _
		                "onSubmit='javascript:return formValidation(this)'"
			                OpenTd ClassCss, "align=left"				               
				                PlaceInput "Eliminar", "submit", "Eliminar", "class=button-OLD"
			                CloseTd
                             PlaceInput "Eliminando", "hidden", "True", ""				
				             PlaceInput "idParametroModificar", "hidden", objRst.Fields("id"), "class=button-OLD"
				             PlaceInput "Parametro", "hidden", Parametro, "class=button-OLD"		                
		                CloseForm
                    CloseTr
                    objRst.MoveNext
                    Loop
                    objRst.Close
                    CloseConnPipelineDB
                    Set adoConn = Nothing
		            if (Contador mod 2) <> 0 then
			            ClassCss = "filaSombra"
		            else
			            ClassCss = "filaBlanca"
		            end if
		            Contador = Contador + 1
		            OpenForm "consi", "post", "segment_general_parameter_edit.asp?Parametro=" _ 
		            & Parametro & "&idParametro=" & idParametro & "&Descripcion=" & Descripcion , _
		            "onSubmit='javascript:return formValidation(this)'"
			            OpenTr ""
				            OpenTd ClassCss, ""
					            PlaceInput "Valor", "TextBox", "" , " class=txboxGenericas"
				            CloseTd
				            OpenTd ClassCss, "align=left "					            
					            PlaceInput "Adicionar", "submit", "Adicionar", "class=button-OLD"
				            CloseTd
                            PlaceInput "Adicionando", "hidden", "True", ""								
					        PlaceInput "Parametro", "hidden", Parametro, "class=button-OLD"
				            OpenTd ClassCss, "align=center "
				            CloseTd
			            CloseTr
		            CloseForm
                CloseTr
            ctbl
        closetd
    closetr
    opentr""
        OpenTd "''", ""
            Response.Write "<br><br>"
		CloseTd
    closetr
    OpenTr ""
		OpenTd "''", ""				
			OpenForm "cons", "post", "segment_general_parameter.asp", "onSubmit='javascript:return formValidation(this)'"
				PlaceInput "Retornar", "submit", "Retornar", "class=button-OLD"
			CloseForm
		CloseTd				
	closetr            
ctbl




%>
	<p/>
        <p/>
		</div>

		<div class="rounded">
			<b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b>
		</div>		
	</body>
</html>
<%
%>
