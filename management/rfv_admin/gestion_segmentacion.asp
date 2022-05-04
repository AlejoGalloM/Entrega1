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
Dim arrClasificacion	' registros dentro de un 
Dim arrClasificacionCalculado ' arreglo con los registros de la segmentacion calculada.
Dim arrParametros	' Parametros del query
Dim objRst			' Recordset object
Dim objTypeDesc		' Recordset object para el typeDescription
Dim strType			' Tipo de parametros puede ser 'Tipo', 'Recencia', 'Frecuencia' o 'Valor'
Dim classid			' clase del control, sirve para almacenar la clase del css
Dim errorCatcher	' variable de estado
Dim Mensaje			' Sirva para mostrar un mensaje en caso de error
Dim EncabezadoParametros ' Contiene el Combo de periodos
Dim Periodo			' parametro del periodo
Dim j				' Contador
Dim conexion
' ***************** POR HACER ****
' Authorize 0,8			AUTORIZACION !!!!!!!!!!!
' *****************
Periodo = Request.Form("Periodo")

Set adoConn = GetConnPipelineDB
' ***************** POR HACER ****
'	DETERMINAR EL LOG !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'	write_sp_log adoConn, 8700, "", 0, "", "", 0, 0, "", "management/default.asp Loaded by " & Session("sp_miLogin")
' ***************** 
'Set adoConn = Server.CreateObject("ADODB.Connection")
'= VERBO EN EJECUCION ===============================================================
On error goto 0


Set objRst = Server.CreateObject("ADODB.Recordset")
'on error goto 0 
strSQL = "exec trfv.dbo.Segmentacion_GetPeriodosCalculados "
objRst.Open strSQL, adoConn, 3
arrParametros = objRst.GetRows()
objRst.Close
set objRst = nothing



Set objRst = Server.CreateObject("ADODB.Recordset")
'on error goto 0 
If Periodo = "" Then
	strSQL = "exec trfv.dbo.SegmentacionGestion_DistribucionSegmentacion "
Else
	strSQL = "exec trfv.dbo.SegmentacionGestion_DistribucionSegmentacion " & Periodo
End If
conexion=strSQL
objRst.Open strSQL, adoConn, 3
arrClasificacion = objRst.GetRows()
objRst.Close
set objRst = nothing


Set objRst = Server.CreateObject("ADODB.Recordset")
'on error goto 0 
If Periodo = "" Then
	strSQL = "exec trfv.dbo.SegmentacionGestion_DistribucionSegmentacion_Calculado "
Else
	strSQL = "exec trfv.dbo.SegmentacionGestion_DistribucionSegmentacion_Calculado " & Periodo
End If
objRst.Open strSQL, adoConn, 3
arrClasificacionCalculado = objRst.GetRows()
objRst.Close
conexion=conexion&" - "&strSQL
set objRst = nothing

CloseConnPipelineDB

write_dataLog Response.Status,"gestion_segmentacion.asp", "gestion_segmentacion.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"exec trfv.dbo.Segmentacion_GetPeriodosCalculados -"&conexion,"N/A","null","Consulta","N/A"

Set adoConn = Nothing

'************ COMBO DE PARAMETROS **************
EncabezadoParametros = "<SELECT name=Periodo class=listagenerica onChange='javascript:changeLocation(this)'>" & vbCrLf 
for J = 0  to UBound(arrParametros,2)
		on error goto 0
'			errorCatcher = arrParametros(0,J)
'			If Err.number <> 0 Then
'				exit for
'			end if
		if trim( Periodo ) = trim( arrParametros(0,J) ) then
			EncabezadoParametros = EncabezadoParametros & _
			"<OPTION SELECTED value='" & arrParametros(0,J) & "'>" & arrParametros(0,J)& "  </OPTION>" & vbCrLf
		else
			EncabezadoParametros = EncabezadoParametros & _
			"<OPTION value='" & arrParametros(0,J) & "'>" & arrParametros(0,J)& "  </OPTION>" & vbCrLf
		end if
Next
If InStr(1,EncabezadoParametros,"SELECTED") = 0 Then
	 Periodo = trim( arrParametros(0,0) )
End If

OpenHTML
OpenHead
'PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
PlaceMeta "Pragma", "", "no_cache"
PlaceLink "REL", "stylesheet", "../../css/OLDMutualStyle.css", "text/css"
%>
<SCRIPT LANGUAGE=javascript src=../../operations/_pipeline_scripts/validation.js></SCRIPT>
</head>
<body class="cuerpo">
	<div class="contenido">
		</br></br>
		<div class="subtituloPagina">
			Segmento por Cliente
		</div>
		</br></br>
<%

otbl"tblcontenido"
    opentr""
        opentd"",""
            Response.Write Mensaje 
	        otbl"tblvalores"
                opentr""
                    OpenTd "''", ""
                        Response.Write "<br><br>"
		            CloseTd
                closetr
                OpenForm "cons", "post", "gestion_segmentacion.asp", "onSubmit='javascript:return formValidation(this)'"
		            OpenTr "valign=middle"
			            OpenTd "'labels'", ""				            
                            Response.Write 	"Parametros Consulta"
			            CloseTd
                        OpenTd "''", ""				            
                            Response.Write 	EncabezadoParametros
			            CloseTd
			            OpenTd "''", ""
					        PlaceInput "Enviar", "submit", "Enviar", "class=button-OLD"
			            CloseTd
		            CloseTr
                CloseForm
	        ctbl
        closetd
    closetr
    opentr""
        opentd"",""
            Response.Write "<br><br>"
        closetd
    closetr
    opentr""
        opentd"",""
            otbl"tblvalores"
                OpenTr ""
				    OpenTd "'titulotabla'", "align=center colspan=8"
					    Response.Write "Información Segmento / Clientes (Asignado)"
				    CloseTd
			    CloseTr
                OpenTr ""
				    OpenTd "'separadorsecciones'", "align=center colspan=8"
				    CloseTd
			    CloseTr
			    OpenTr ""
				    OpenTd "'texto-informativo'", "align=center"					   
				    CloseTd
			    CloseTr
                OpenTr ""
			        OpenTh "''", "align=center "
				        Response.Write "Segmento"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write " "
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Característica"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Nro. de Clientes"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "% participación clientes"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Inf. Básica"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Top Sociedades"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Top FPs"
			        CloseTh
		        CloseTr
                Dim LastSegment
                LastSegment = ""
                for J = 0  to UBound(arrClasificacion,2)
		            on error resume next
			        errorCatcher = arrClasificacion(0,J)
			        If Err.number <> 0 Then
				        exit for
			        end if
		            if (2 * Round(J / 2)) = J then
			            classid = "filaSombra"
		            else
			            classid = "filaBlanca"
		            end if
		            If isnull( arrClasificacion(1,J) ) Then
			            classid = "filaSombra"
		            End if		
		            OpenTr ""
		                OpenForm "cons", "post", "gestionsegmentacion.asp", "onSubmit='javascript:return formValidation(this)'"
			                OpenTd classid, "align=center "
				                If LastSegment <> arrClasificacion(0,J) Then
					                Response.Write arrClasificacion(0,J)
					                LastSegment = arrClasificacion(0,J)
				                Else
					                Response.Write " " 
				                End If
			                CloseTd
			                OpenTd classid, "align=center "
				                If isnull( arrClasificacion(1,J) ) Then
					                Response.Write "Total"
				                Else
					                Response.Write " "
				                End If
			                CloseTd
			                OpenTd classid, "align=center "
				                If Not isnull( arrClasificacion(1,J) ) Then
					                Response.Write arrClasificacion(1,J)
				                Else
					                Response.Write " "
				                End If
			                CloseTd
			                OpenTd classid, "align=center "
				                If IsNull( arrClasificacion(1,J) ) Then
					                Response.Write "<bold>"
				                End If
				                Response.Write FormatNumber( arrClasificacion(2,J) ,0)
				                If IsNull( arrClasificacion(1,J) ) Then
					                Response.Write "</bold>"
				                End If
			                CloseTd			
			                OpenTd classid, "align=center "
				                If IsNull( arrClasificacion(1,J) ) Then
					                Response.Write "<bold>"
				                End If
                On error goto 0				
					                Response.Write FormatPercent( CDbl( arrClasificacion(3,J)) )
				                If IsNull( arrClasificacion(1,J) ) Then
					                Response.Write "</bold>"
				                End If
			                CloseTd
			                OpenTd classid, "align=center "
				                If isnull( arrClasificacion(1,J) ) And Not isnull( arrClasificacion(0,J) )   Then					
					                enlacetabla "gestion_informacionclientessegmento.asp?periodo=" & Periodo & "&segmento=" & arrClasificacion(0,J) , "Info Básica", "class=enlacetabla"
				                Else
					                Response.Write " "
				                End If				
			                CloseTd
			                OpenTd classid, "align=center "
				                If isnull( arrClasificacion(1,J) ) And Not isnull( arrClasificacion(0,J) )   Then
					                enlacetabla "gestion_informaciontop20sociedadsegmento.asp?periodo=" & Periodo & "&segmento=" & arrClasificacion(0,J) , "Info Básica", "class=enlacetabla"				
				                Else
				                End If				
			                CloseTd
			                OpenTd classid, "align=center "
				                If isnull( arrClasificacion(1,J) ) And Not isnull( arrClasificacion(0,J) )  Then
					                enlacetabla "gestion_informaciontop20agentessegmento.asp?periodo=" & Periodo & "&segmento=" & arrClasificacion(0,J) , "Info Básica", "class=enlacetabla"				
				                Else
					                Response.Write " "
				                End If				
			                CloseTd			
		                CloseForm
		        CloseTr
                next
                OpenTr ""
			        OpenTd "", "align=center "				
				        OpenForm "cons", "post", "gestionsegmentacion.asp", "onSubmit='javascript:return formValidation(this)'"
				        CloseForm
			        CloseTd				
		        CloseTr
            ctbl
        closetd
    closetr
    opentr""
        opentd"",""
            Response.Write "<br><br>"
        closetd
    closetr
    opentr""
        opentd"",""
            otbl"tblvalores"
	            openTr ""
				    OpenTd "'titulotabla'", "align=center colspan=8"
					    Response.Write "Información Segmento / Clientes (Calculado)"
				    CloseTd
			    CloseTr
                OpenTr ""
				    OpenTd "'separadorsecciones'", "align=center colspan=8"
				    CloseTd
			    CloseTr
			    OpenTr ""
				    OpenTd "'texto-informativo'", "align=center"					   
				    CloseTd
			    CloseTr
                OpenTr ""
			        OpenTh "''", "align=center "
				        Response.Write "Segmento"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write " "
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Característica"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "Nro. de Clientes"
			        CloseTh
			        OpenTh "''", "align=center "
				        Response.Write "% participación clientes"
			        CloseTh
		        CloseTr
                LastSegment = ""
                for J = 0  to UBound(arrClasificacionCalculado,2)
		            on error resume next
			        errorCatcher = arrClasificacionCalculado(0,J)
			        If Err.number <> 0 Then
				        exit for
			        end if
		            if (2 * Round(J / 2)) = J then
			            classid = "filaSombra"
		            else
			            classid = "filaBlanca"
		            end if
		            If isnull( arrClasificacionCalculado(1,J) ) Then
			            classid = "filaSombra"
		            End if		
		            OpenTr ""
		                OpenForm "cons", "post", "gestionsegmentacion.asp", "onSubmit='javascript:return formValidation(this)'"
			                OpenTd classid, "align=center "
				                If LastSegment <> arrClasificacionCalculado(0,J) Then
					                Response.Write arrClasificacionCalculado(0,J)
					                LastSegment = arrClasificacionCalculado(0,J)
				                Else
					                Response.Write " " 
				                End If
			                CloseTd
			                OpenTd classid, "align=center "
				                If isnull( arrClasificacionCalculado(1,J) ) Then
					                Response.Write "Total"
				                Else
					                Response.Write " "
				                End If
			                CloseTd

			                OpenTd classid, "align=center "
				                If Not isnull( arrClasificacionCalculado(1,J) ) Then
					                Response.Write arrClasificacionCalculado(1,J)
				                Else
					                Response.Write " "
				                End If
			                CloseTd
			                OpenTd classid, "align=center "
				                If IsNull( arrClasificacionCalculado(1,J) ) Then
					                Response.Write "<bold>"
				                End If
				                Response.Write FormatNumber( arrClasificacionCalculado(2,J) ,0)
				                If IsNull( arrClasificacionCalculado(1,J) ) Then
					                Response.Write "</bold>"
				                End If
			                CloseTd
			                OpenTd classid, "align=center "
				                If IsNull( arrClasificacionCalculado(1,J) ) Then
					                Response.Write "<bold>"
				                End If
                On error goto 0				
				                Response.Write FormatPercent( CDbl( arrClasificacionCalculado(3,J)) )
				                If IsNull( arrClasificacionCalculado(1,J) ) Then
					                Response.Write "</bold>"
				                End If
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
		OpenTd "", "align=center "				
			OpenForm "cons", "post", "gestionsegmentacion.asp", "onSubmit='javascript:return formValidation(this)'"
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
