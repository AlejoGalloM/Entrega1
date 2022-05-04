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
'File Name:		segmentacionprioridad.asp xxyy
'Path:				management/rfv_admin/ segmentacionprioridad.asp
'Created By:		G. Pinerez 2004/04/26
'Last Modified:	
'					Mar�a Margarita Cardozo 2004/11/23 Se adicionaron campos y combos
'						
'Modifications:	
'Parameters:		
'				
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
Dim TypeDescription ' Descripcion del tipo
Dim Mensaje			' Sirvae para mostrar un mensaje en caso de error
Dim SegVIP
Dim Id_SegmentoClasificacion 'Campo SegmentoClasificacion
Dim Id_SegmentoCaracteristica 'Campo SegmentoCaracteristica
dim CmbSegmentoClasificacion 'Combo Segmento Clasificacion
dim CmbSegmentoCaracteristica ' Combo Segmento Caracteristica
Dim arrSegmentoClasificacion	' Cargar SegmentoClasificacion
Dim arrSegmentoCaracteristica	' Cargar SegmentoCaracteristica

'###########3 PARAMETROS DE LA TABLA SEGMENTACION 
Dim Par_Segmento
Dim Par_TipoProcedimiento
Dim Par_Descripcion
Dim Par_Caracteristica
Dim Par_Prioridad
Dim Par_Estilo
Dim Par_BajaAuto
Dim Par_Defecto
Dim Par_SegVIP
Dim Par_Id_SegmentoClasificacion
Dim Par_Id_SegmentoCaracteristica

'###########3 LOGS
Dim processInfo, component_id
Dim valueLog

'###########3 PARAMETROS DE LA TABLA SEGMENTACION 
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



Par_Segmento = trim( Request.Form("Par_Segmento") )
Par_TipoProcedimiento = trim( Request.Form("Par_TipoProcedimiento") )
Par_Descripcion = trim( Request.Form("Par_Descripcion") )
Par_Caracteristica = trim( Request.Form("Par_Caracteristica") )
Par_Prioridad = trim( Request.Form("Par_Prioridad") )
Par_Estilo = trim( Request.Form("Par_Estilo") )
Par_Defecto = Request.Form("Par_Defecto") 

Par_SegVIP = Request.Form("Par_SegVIP")
Par_Id_SegmentoClasificacion = Request.Form("Par_Id_SegmentoClasificacion") 
Par_Id_SegmentoCaracteristica = Request.Form("Par_Id_SegmentoCaracteristica") 


if trim( Request.Form("Par_BajaAuto") ) = 1 then
	Par_BajaAuto = true
else
	Par_BajaAuto = false
end if
if trim( Request.Form("Par_Defecto") ) = "True" then 
	Par_Defecto = true
else
	Par_Defecto = false
end if

component_id = "segmentacionprioridad.asp"
processInfo =  "management/default.asp Loaded by " & Session("sp_miLogin")

' ***************** POR HACER ****
' Authorize 0,8			AUTORIZACION !!!!!!!!!!!
' *****************

Set adoConn = GetConnPipelineDB
' ***************** POR HACER ****
'	DETERMINAR EL LOG !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'	write_sp_log adoConn, 8700, "", 0, "", "", 0, 0, "", "management/default.asp Loaded by " & Session("sp_miLogin")
' ***************** 
'Set adoConn = Server.CreateObject("ADODB.Connection")
'= VERBO EN EJECUCION ===============================================================
On error goto 0
if Modificando = "True" then	
'	adoConn.Execute "exec trfv..segmentacion_BorrarSegmento " & Borrar 
' Cambio mmc elimino esta linea y adiciono a la Insercion que borre en caso de que exista
' ya que si el usuario selecciona modificar y no presina actualizar se borra el registro

	Session("valueLogOld")="Par_Segmento: "&Par_Segmento&"','"&"Par_TipoProcedimiento: "&Par_TipoProcedimiento&"',"&"Par_Id_SegmentoClasificacion: "&Par_Id_SegmentoClasificacion&","&"Par_Id_SegmentoCaracteristica: "&Par_Id_SegmentoCaracteristica&","&"Par_Prioridad: "&Par_Prioridad&",'"&"Par_Estilo: "&Par_Estilo& "','"&"User: "&Session("sp_milogin")&"', 1"& ", "&"BitValue: "&BitValue&","&"BitValueSegVIP: "&BitValueSegVIP
	write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "trfv..segmentacion_BorrarSegmento " ,Session.contents("valueLogOld"),"","Operación-Modificación","N/A"
	
end if
if Borrando = "True" then	
 	Session("valueLogOld")="Par_Segmento: "&Par_Segmento&"','"&"Par_TipoProcedimiento: "&Par_TipoProcedimiento&"',"&"Par_Id_SegmentoClasificacion: "&Par_Id_SegmentoClasificacion&","&"Par_Id_SegmentoCaracteristica: "&Par_Id_SegmentoCaracteristica&","&"Par_Prioridad: "&Par_Prioridad&",'"&"Par_Estilo: "&Par_Estilo& "','"&"User: "&Session("sp_milogin")&"', 1"& ", "&"BitValue: "&BitValue&","&"BitValueSegVIP: "&BitValueSegVIP
	adoConn.Execute "exec trfv..segmentacion_BorrarSegmento " & Borrar
	write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "trfv..segmentacion_BorrarSegmento " ,Session.contents("valueLogOld"),"","Operación-Modificación","N/A" 
end if
if Adicionando = "True" then	

	if trim(Par_Id_SegmentoClasificacion)="" then
		msgbox("Por favor seleccione un SegmentoClasificacion")
	end if
		
	if trim(Par_Id_SegmentoCaracteristica)="" then
		msgbox("Por favor seleccione un SegmentoCaracteristica")
	end if

	valueLog = "Par_Segmento: "&Par_Segmento&"','"&"Par_TipoProcedimiento: "&Par_TipoProcedimiento&"',"&"Par_Id_SegmentoClasificacion: "&Par_Id_SegmentoClasificacion&","&"Par_Id_SegmentoCaracteristica: "&Par_Id_SegmentoCaracteristica&","&"Par_Prioridad: "&Par_Prioridad&",'"&"Par_Estilo: "&Par_Estilo& "','"&"User: "&Session("sp_milogin")&"', 1"& ", "&"BitValue: "&BitValue&","&"BitValueSegVIP: "&BitValueSegVIP
	
	write_dataLog Response.Status, component_id,processInfo,Session.contents("name"), "exec trfv..segmentacion_insertarSegmento ",Session.contents("valueLogOld"),valueLog,"Operación-Adición","N/A"
	
	
	Dim BitValue
	Dim BitValueSegVIP
	If Request("Par_Defecto") then 
		BitValue = 1
	Else
		BitValue = 0 
	End if
	
	If Request("Par_SegVIP") then 
		BitValueSegVIP = 1
	Else
		BitValueSegVIP = 0 
	End if
	
	
	strSQL= "exec trfv..segmentacion_insertarSegmento " & _
		"'" & Par_Segmento & "','" & _
		Par_TipoProcedimiento  & _
		 "',"  & Par_Id_SegmentoClasificacion  &   _
		 ","  & Par_Id_SegmentoCaracteristica &   _
		 ","  & Par_Prioridad   & _
		 ",'" & Par_Estilo  & "','" & Session("sp_milogin") & "', 1" &  _
		 ", " & BitValue &  _
		 "," & BitValueSegVIP  

	adoConn.Execute strSQL

	If err.number <> 0 then		
		if err.Description = "[Microsoft][ODBC SQL Server Driver][SQL Server]INSERT statement conflicted with " & _
			"COLUMN FOREIGN KEY constraint 'FK_Segmentacion_TipoProcedimiento'. The conflict " & _
			"occurred in database 'TRFV', table 'TipoProcedimiento', column " & _
			"'TipoProcedimiento'." then
			Response.Write "<SCRIPT>window.alert('El procedimiento no existe en la base de datos');</SCRIPT>" & vbCrLf
		End if
	End if
	

Par_Segmento = ""
Par_TipoProcedimiento = ""
Par_Descripcion = ""
Par_Caracteristica = ""
Par_Prioridad = ""
Par_Estilo = ""
Par_Defecto = ""
Par_BajaAuto = ""
Par_SegVIP =""
Par_Id_SegmentoCaracteristica=""
Par_Id_SegmentoClasificacion=""
		
end if
'= VERBO EN EJECUCION ===============================================================

Set objRst = Server.CreateObject("ADODB.Recordset")
'on error goto 0 
strSQL = "exec trfv.dbo.segmentacion_GetPrioridades "
objRst.Open strSQL, adoConn, 3
arrParameters = objRst.GetRows()
objRst.Close

'POBLAR COMBOS SEGMENTOCLASIFICACION SEGMENTO CARACTERISTICA CAMBIO MMC 2004/11/16
strSQL = "exec trfv.dbo.segmentacion_GetSegmentoClasificacion "
objRst.Open strSQL, adoConn, 3
arrSegmentoClasificacion = objRst.GetRows()
objRst.Close

strSQL = "exec trfv.dbo.segmentacion_GetSegmentoCaracteristica "
objRst.Open strSQL, adoConn, 3
arrSegmentoCaracteristica = objRst.GetRows()
objRst.Close

'FIN CAMBIO

set objRst = nothing
CloseConnPipelineDB

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
			Parametros de segmentaci�n
		</div>
		</br></br>
<%
Response.Write Mensaje 

otbl"tblcontenido"
    opentr""
        opentd"",""
            otbl"tblvalores"
                opentr""
                    OpenTd "''", ""
                        Response.Write "<br>"
		            CloseTd
                closetr
		        OpenTr ""
			        OpenTh "''", "align=center "
				        Response.Write "Segmento"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Tipo Proced."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Segmento descr."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Caracter�st."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Prior."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Estilo"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Por defecto"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Seg VIP"
			        CloseTh
				    OpenTh "''", "align=center "
				        Response.Write ""
			        CloseTh
				    OpenTh "''", "align=center "
				        Response.Write ""
                    closeth
		        CloseTr		
		        OpenTr ""
		            OpenForm "valores", "post", "segmentacionprioridad.asp", "onSubmit='javascript:return formValidation(this)'"		
			            OpenTd "''", "align=center "
				            PlaceInput "Par_Segmento", "Text", Par_Segmento, "id='R' size=6 class=txboxGenericas"
			            CloseTd		
			            OpenTd "''", "align=center "
				            PlaceInput "Par_TipoProcedimiento", "Text", Par_TipoProcedimiento, "id='R' size=6, class=txboxGenericas"
			            CloseTd
			            OpenTd "''", "align=center "
					        If IsArray(arrSegmentoClasificacion) Then
						        CmbSegmentoClasificacion = "<SELECT name=Par_Id_SegmentoClasificacion class=listagenerica id='R              CmbSegmentoClasificacion'>" & vbCrLf & _
						        "<OPTION value=''>Seg Clasif</OPTION>" & vbCrLf
						        For J = 0 To UBound(arrSegmentoClasificacion, 2)'rows		
								        if ( Request.Form("Par_Id_SegmentoClasificacion")=cstr(arrSegmentoClasificacion(3,J))) then
									        CmbSegmentoClasificacion = CmbSegmentoClasificacion & _
									        "<OPTION  SELECTED value='" & arrSegmentoClasificacion(3,J) & "'>" & _
									        arrSegmentoClasificacion(0,J) & _ 				
									        "</OPTION>" & vbCrLf	
								        else						
									        CmbSegmentoClasificacion = CmbSegmentoClasificacion & _
									        "<OPTION value='" & arrSegmentoClasificacion(3,J) & "'>" & _
									        arrSegmentoClasificacion(0,J) & _ 				
									        "</OPTION>" & vbCrLf	
								        end if							
						        next						
						        CmbSegmentoClasificacion = CmbSegmentoClasificacion & "</SELECT>"
						        response.write CmbSegmentoClasificacion
					        else 
						        response.write "No hay parametros creados"
					        end if
			            CloseTd
			            OpenTd "''", "align=center "
			                If IsArray(arrSegmentoCaracteristica) Then
				                CmbSegmentoCaracteristica = "<SELECT name=Par_Id_SegmentoCaracteristica class=listagenerica id='R              CmbSegmentoCaracteristica'>" & vbCrLf & _
				                "<OPTION value=''>Seg Clasif</OPTION>" & vbCrLf
				                For J = 0 To UBound(arrSegmentoCaracteristica, 2)'rows				
					                if ( Request.Form("Par_Id_SegmentoCaracteristica")=cstr(arrSegmentoCaracteristica(0,J))) then
						                CmbSegmentoCaracteristica = CmbSegmentoCaracteristica & _
						                "<OPTION  SELECTED value='" & arrSegmentoCaracteristica(0,J) & "'>" & _
						                arrSegmentoCaracteristica(1,J) & _ 				
						                "</OPTION>" & vbCrLf	
					                else			
						                CmbSegmentoCaracteristica = CmbSegmentoCaracteristica & _
						                "<OPTION value='" & arrSegmentoCaracteristica(0,J) & "'>" & _
						                arrSegmentoCaracteristica(1,J) & _ 				
						                "</OPTION>" & vbCrLf	
					                end if	
				                next
				                CmbSegmentoCaracteristica = CmbSegmentoCaracteristica & "</SELECT>"
				                response.write CmbSegmentoCaracteristica
			                else 
				                response.write "No hay parametros creados"
			                end if
			            CloseTd			
			            OpenTd "teven", "align=center "
				            PlaceInput "Par_Prioridad", "Text", Par_Prioridad, "id='R' size=2 class=txboxGenericas"
			            CloseTd
			            OpenTd "teven", "align=center "
				            PlaceInput "Par_Estilo", "Text", Par_Estilo, "id='R' size=8 class=txboxGenericas"
			            CloseTd
			            OpenTd "teven", "align=center "
				            If Par_Defecto = "True" then
					            PlaceInput "Par_Defecto", "checkbox", "true" , " CHECKED id='R' onClick='javascript:Par_Defecto.value=!Par_Defecto.checked; '"
				            Else
					            PlaceInput "Par_Defecto", "checkbox", "false" , "  id='R'  onClick='javascript:Par_Defecto.value=Par_Defecto.checked; '"
				            End If				
			            CloseTd		
			            OpenTd "teven", "align=center "
				            If Par_Defecto = "True" then
					            PlaceInput "Par_SegVIP", "checkbox", "true" , " CHECKED id='R' onClick='javascript:Par_SegVIP.value=!Par_SegVIP.checked; '"
				            Else
					            PlaceInput "Par_SegVIP", "checkbox", "false" , "  id='R'  onClick='javascript:Par_SegVIP.value=Par_SegVIP.checked; '"
				            End If				
			            CloseTd
                        opentd"''","width=20px"
                        closetd		
			            OpenTd "teven", "align=center "
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
                opentr""
                    OpenTd "''", ""
                        Response.Write "<br><br>"
		            CloseTd
                closetr
            ctbl
        closetd
    closetr
    opentr""
        opentd"",""
            otbl"tblvalores"
		        OpenTr ""
			        OpenTh "''", "align=center "
				        Response.Write "Segmento"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Tipo Proced."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Segmento<br>Descrp."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Caracter�stica"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Prior."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Estilo"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "por defecto"
			        CloseTh			
			        OpenTh "''", "align=center "
				        Response.Write "Seg VIP"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Id Seg Clas."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Id Seg Car."
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write ""
			        CloseTh
		        CloseTr
                for J = 0  to UBound(arrParameters,2)
		        on error resume next
			        errorCatcher = arrParameters(0,J)
			        If Err.number <> 0 Then
				        exit for
			        end if			
		        if (2 * Round(J / 2)) = J then
			        classid = "filaSombra align=center"
		        else
			        classid = "filaBlanca align=center"
		        end if
		        OpenTr ""
                    OpenForm "cons", "post", "segmentacionprioridad.asp", "onSubmit='javascript:return formValidation(this)'"
			            OpenTd classid, "align=center "
				            Response.Write arrParameters(0,J)
			            CloseTd
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
				            Response.Write arrParameters(4,J)
			            CloseTd
			            OpenTd classid, "align=center "
				            Response.Write arrParameters(7,J)
			            CloseTd
			            OpenTd classid, "align=center "
				            If arrParameters(9,J) = "True" then				
					            Response.Write "<input disabled name=Par_Defectox type=checkbox value='' CHECKED>"
				            Else
					            Response.Write "<input disabled name=Par_Defectox type=checkbox value='' >"
				            End If
			            CloseTd	
			            OpenTd classid, "align=center "
				            If arrParameters(10,J) = "True" then				
					            Response.Write "<input disabled name=Par_SegVIPx type=checkbox value='' CHECKED>"
				            Else
					            Response.Write "<input disabled name=Par_SegVIPx type=checkbox value='' >"
				            End If
			            CloseTd
			            OpenTd classid, "align=center "
					            Response.Write arrParameters(11,J)
			            CloseTd			
			            OpenTd classid, "align=center "
					            Response.Write arrParameters(12,J)
			            CloseTd
			            OpenTd classid, "align=center "		
				            PlaceInput "Par_Segmento", "hidden", arrParameters(0,J) , ""
				            PlaceInput "Par_TipoProcedimiento", "hidden", arrParameters(1,J) , ""
				            PlaceInput "Par_Descripcion", "hidden", arrParameters(2,J) , ""				
				            PlaceInput "Par_Caracteristica", "hidden", arrParameters(3,J) , ""
				            PlaceInput "Par_Prioridad", "hidden", arrParameters(4,J) , ""
				            PlaceInput "Par_Estilo", "hidden", arrParameters(7,J) , ""				
				            PlaceInput "Par_BajaAuto", "hidden", arrParameters(8,J) , ""				
				            PlaceInput "Par_Defecto", "hidden", arrParameters(9,J) , ""	
				            PlaceInput "Par_SegVIP", "hidden", arrParameters(10,J) , ""				
				            PlaceInput "Par_Id_SegmentoClasificacion", "hidden", arrParameters(11,J) , ""				
				            PlaceInput "Par_Id_SegmentoCaracteristica", "hidden", arrParameters(12,J) , ""
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
                closetr
                NEXT
            ctbl
        closetd
    closetr
    opentr""
        OpenTd "''", ""
            Response.Write "<br>"
		CloseTd
    closetr
    opentr""
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
<%
%>
