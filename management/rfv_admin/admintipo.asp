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
'File Name:		admintipo.asp xxyy
'Path:				management/rfv_admin
'Created By:		G. Pinerez 2002/02/26
'Last Modified:	
'						
'Modifications:	
'Parameters:		tipo : tipo de parametro, puede ser 'Tipo', 
'					'Recencia', 'Frecuencia' o 'Valor'
'Returns:			Mangement default page
'Additional Information:
'ToDoes : "POR HACER ****"
'===================================================================================
On Error Resume Next
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1

'===================================================================================
Dim adoConn			' conexion a la BD
Dim strSQL			' string que contiene el query
Dim arrParameters	' registros dentro de un 
Dim objRst			' Recordset object
Dim objTypeDesc		' Recordset object para el typeDescription
Dim strType			' Tipo de parametros puede ser 'Tipo', 'Recencia', 'Frecuencia' o 'Valor'
Dim J				' Contador
Dim classid			' clase del control, sirve para almacenar la clase del css
Dim errorCatcher	' variable de estado
Dim Borrando		' verbo, determina si accion que se esta ejecutando es borrar
Dim Borrar			' cual registro borrar
Dim Adicionando		' verbo, determina si accion que se esta ejecutando es borrar
Dim Modificando		' determina si accion que se esta ejecutando es borrar
Dim Modificar			' cual registro Modificar
Dim Clasificacion	' campo de clasificacion
Dim RInferior		' campo de rango inferior
Dim RSuperior		' campo de rango superior
Dim Clasificacionm	' campo de clasificacion de modificacion
Dim RInferiorm		' campo de rango inferior de modificacion
Dim RSuperiorm		' campo de rango superior de modificacion
Dim TypeDescription ' Descripcion del tip
Dim valueNewLog		'Valores nuevos para micoservicio de log'
Dim valueLog		'Valores nuevos para micoservicio de log'
Dim processInfo, component_id
'===================================================================================


strType = Request.Form("Tipo")
Borrando = Request.Form("Borrando")
Borrar = Request.Form("Borrar")

Adicionando = Request.Form("Adicionando")
Clasificacion = Request.Form("Clasificacion")
RInferior = Request.Form("RInferior")
RSuperior = Request.Form("RSuperior")

Clasificacionm = Request.Form("Clasificacionm")
RInferiorm = Request.Form("RInferiorm")
RSuperiorm = Request.Form("RSuperiorm")
Modificando = Request.Form("Modificando")

' ***************** POR HACER ****
' Authorize 0,8			AUTORIZACION !!!!!!!!!!!
' *****************


component_id = "admintipo.asp"
processInfo =  "management/default.asp Loaded by " & Session("sp_miLogin")

Set adoConn = GetConnPipelineDB
' ***************** POR HACER ****
'	DETERMINAR EL LOG !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'	write_sp_log adoConn, 8700, "", 0, "", "", 0, 0, "", "management/default.asp Loaded by " & Session("sp_miLogin")
' ***************** 
'Set adoConn = Server.CreateObject("ADODB.Connection")
'= VERBO EN EJECUCION ===============================================================
if Modificando = "True" then	
	adoConn.Execute "exec trfv..sprfv_deleteParameter " & Borrar 

	Session("valueLog") = "Clasificacion: "&Clasificacionm&","&"RInferior:"&RInferiorm&","&"RSuperior: "&RSuperiorm
	write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "trfv..sprfv_deleteParameter - trfv.dbo.sprfv_getParameters - trfv.dbo.sprfv_getTypeDescription'" & Trim(strType) & "'"&"exec trfv.dbo.sprfv_getTypeDescription  '" & Trim(strType) & "'" ,myStr,valueNewLog,"Operación-Modificación","N/A"

end if
if Borrando = "True" then	
	adoConn.Execute "exec trfv..sprfv_deleteParameter " & Borrar 
	
	Session("valueLog") = "Clasificacion: "&Clasificacionm&","&"RInferior:"&RInferiorm&","&"RSuperior: "&RSuperiorm
	write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "trfv..sprfv_deleteParameter - trfv.dbo.sprfv_getParameters - trfv.dbo.sprfv_getTypeDescription'" & Trim(strType) & "'"&"exec trfv.dbo.sprfv_getTypeDescription  '" & Trim(strType) & "'" ,Session.contents("valueLog"),valueNewLog,"Operación-Eliminación","N/A"


end if
if Adicionando = "True" then	
	adoConn.Execute "exec trfv..sprfv_addParameter '" & Trim(strType) & "',"  & _
		Clasificacion & "," & RInferior & "," &  RSuperior

		valueNewLog = "Clasificacion: "&Clasificacion&","&"RInferior:"&RInferior&","&"RSuperior: "&RSuperior

		
		write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "sprfv_addParameter - trfv.dbo.sprfv_getParameters - trfv.dbo.sprfv_getTypeDescription '" & Trim(strType) & "'"&"exec trfv.dbo.sprfv_getTypeDescription  '" & Trim(strType) & "'" ,Session.contents("valueLog"),valueNewLog,"Operación-Adicionar","N/A"

end if
'= VERBO EN EJECUCION ===============================================================

Set objRst = Server.CreateObject("ADODB.Recordset")
'on error goto 0 
strSQL = "exec trfv.dbo.sprfv_getParameters '" & Trim(strType) & "'"
objRst.Open strSQL, adoConn, 3
arrParameters = objRst.GetRows()
objRst.Close
set objRst = nothing

Set objTypeDesc = Server.CreateObject("ADODB.Recordset")
'on error goto 0 
strSQL = "exec trfv.dbo.sprfv_getTypeDescription  '" & Trim(strType) & "'"
objTypeDesc.Open strSQL, adoConn, 3
TypeDescription = objTypeDesc.Fields("ParameterDescription")

objTypeDesc.Close
CloseConnPipelineDB
set objTypeDesc = nothing
Set adoConn = Nothing

OpenHTML
OpenHead
PlaceMeta "Pragma", "", "no_cache"
PlaceLink "REL", "stylesheet", "../../css/OLDMutualStyle.css", "text/css"
%>
<SCRIPT LANGUAGE=javascript src=../../operations/_pipeline_scripts/validation.js></SCRIPT>
</head>
<body class="cuerpo">
	<div class="contenido">
		</br></br>
		<div class="subtituloPagina">
        <%
            Response.Write "Par�metros para " & strType
        %>            
		</div>
		</br></br>
<%

otbl"tblcontenido"
    opentr""
        opentd"",""
            Response.Write "<br>"
        closetd
    closetr
    opentr""
        opentd"''",""
            otbl"tblvalores"
                OpenTr ""
					OpenTd "'tituloTabla'", "colspan=4 align=center"
						 Response.Write TypeDescription
					CloseTd
				CloseTr
                OpenTr ""
					OpenTd "'separadorSecciones'", "colspan=4 align=center"
					CloseTd
				CloseTr
		        OpenTr ""
			        OpenTh "thead", "align=center "
				        Response.Write "Clasificaci�n"
			        CloseTh
			        OpenTh "thead", "align=center "
				        Response.Write "Rango Inferior"
			        CloseTh
			        OpenTh "thead", "align=center "
				        Response.Write "Rango Superior"
			        CloseTh
			        OpenTh "thead", "align=center "
				        Response.Write ""
			        CloseTh
		        CloseTr
		        OpenTr ""
		            OpenForm "valores", "post", "admintipo.asp", "onSubmit='javascript:return formValidation(this)'"
			            OpenTd "''", "align=center "
				            PlaceInput "Clasificacion", "Text", Clasificacionm, "id='R' class=txboxGenericas"
			            CloseTd
			            OpenTd "''", "align=center "
				            PlaceInput "RInferior", "Text", RInferiorm, "id='R' class=txboxGenericas"
			            CloseTd
			            OpenTd "''", "align=center "
				            PlaceInput "RSuperior", "Text", RSuperiorm, "id='R' class=txboxGenericas"
				            Response.Write "Infinito"
				            PlaceInput "Infinito" , "checkbox", "Infinito", "onclick='Javascript:if(RSuperior.value==-1){RSuperior.value="""";}else{RSuperior.value=-1;}'"
			            CloseTd
			            OpenTd "''", "align=center "
				            PlaceInput "Tipo", "hidden", strType , ""
				            PlaceInput "Adicionando", "hidden", "True", ""
                            if Modificando = "True" then	
				               PlaceInput "Adicionar", "submit", "Modificar", "class=button-OLD"
                            else
				               PlaceInput "Adicionar", "submit", "Adicionar", "class=button-OLD"
                            end if
			            CloseTd
		            CloseForm
		        CloseTr
            ctbl
        closetd
    closetr
    opentr""
        opentd"",""
            Response.Write "<br>"
        closetd
    closetr
    opentr""
        opentd"",""
            otbl"tblvalores"
	            OpenTr ""
		            OpenTh "''", "align=center "
			            Response.Write "Par�metro"
		            CloseTh
		            OpenTh "''", "align=center "
			            Response.Write "Clasificaci�n"
		            CloseTh
		            OpenTh "''", "align=center "
			            Response.Write "Rango Inferior"
		            CloseTh
		            OpenTh "''", "align=center "
			            Response.Write "Rango Superior"
		            CloseTh
	            CloseTr
                for J = 0  to UBound(arrParameters,2)
		            on error resume next
			            errorCatcher = arrParameters(0,J)
			            If Err.number <> 0 Then
				            exit for
			            end if
			
		            if (2 * Round(J / 2)) = J then
			            classid = "filaSombra"
		            else
			            classid = "filaBlanca"
		            end if
		            OpenTr ""
		                OpenForm "cons", "post", "admintipo.asp", "onSubmit='javascript:return formValidation(this)'"
			                OpenTd classid, "align=center "
				                Response.Write arrParameters(1,J)
			                CloseTd
			                OpenTd classid, "align=center "
				                Response.Write arrParameters(2,J)
			                CloseTd
			                OpenTd classid, "align=center "
				                Response.Write arrParameters(3,J)
			                CloseTd
			                OpenTd classid, "align=center "
				                if arrParameters(4,J) = "-1" then
					                Response.Write "* infinito * "
				                else
					                Response.Write arrParameters(4,J)
				                end if
			                CloseTd
			                OpenTd classid, "align=center "		
				                PlaceInput "Clasificacionm", "hidden", arrParameters(2,J) , ""
				                PlaceInput "RInferiorm", "hidden", arrParameters(3,J) , ""
				                PlaceInput "RSuperiorm", "hidden", arrParameters(4,J) , ""				
				                PlaceInput "Modificando", "hidden", "False", ""
				                PlaceInput "Modificar", "hidden", arrParameters(0,J), ""
				                PlaceInput "Modificacion", "button", "Modificar", "class=button-OLD onclick='javascript:form.Modificando.value=""True"";form.submit();'"
			                CloseTd				
			                OpenTd classid, "align=center "				
				                PlaceInput "Tipo", "hidden", strType , ""
				                PlaceInput "Borrando", "hidden", "False", ""
				                PlaceInput "Borrar", "hidden", arrParameters(0,J), ""
				                PlaceInput "Borrar", "submit", "Borrar", "class=button-OLD onclick='javascript:form.Borrando.value=""True"";form.submit();'"
			                CloseTd				
		                CloseForm
		            CloseTr
                next
            ctbl
        closetd
    closetr
    opentr""
        opentd"",""
            Response.Write "<br><br>"
        closetd
    closetr
    OpenTr ""
		OpenTd "''", "align=center "				
			OpenForm "cons", "post", "default.asp", "onSubmit='javascript:return formValidation(this)'"
			PlaceInput "Retornar", "submit", "Retornar", "class=button-OLD"
			CloseForm
		CloseTd				
	CloseTr
ctbl
%>
	<p/>
        <p/>
		</div>
	</body>
</html>
