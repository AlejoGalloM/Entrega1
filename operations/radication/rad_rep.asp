<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		rad_rep.asp 10500
'Path:			/operations/radication
'Created By:		Guillermo Aristizabal 2001/09/12
'Last Modified:		A. Orozco 2001/10/08
'			Guillermo Aristizabal 2001/10/11
'                       Rafael Lagos  2002/01/14
'                       Armando Arias - 2008/May/09 - Cumplimiento circular SF052 - PlaceTitle - write_sp_log 13291
'Modifications:
'Parameters:						
'Returns:						
'Additional Information:	
'===================================================================================
Option Explicit
On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="radicationqueries.asp"-->
<%
Authorize 1,14
Response.Write "<link rel='stylesheet' href='../../css/OLDMutualStyle.css' type='text/css'>" & vbCrLf
'== declares ===
Dim Inicio
Dim Fin
Dim FinOld
Dim Section
Dim Cn
Dim rs
Dim sql
Dim DocTypeCombo

'== initials asignments ==
set cn = getconnpipelinedb
set rs = Server.CreateObject("ADODB.RecordSet")

DocTypeCombo = PlaceDocTypeCombo ("listagenerica", cn, "C")


sql = "sprd_get_RadUserSection " & Session.Contents("sp_idworker")
rs.Open sql,cn,3

if rs.BOF and rs.EOF then
	Section = "-1"
else
  Section = rs.Fields(0)
end if
rs.Close


write_sp_log cn, 13291, "sprd_get_RadUserSection", 0, "", "", 0, 0, "", "rad_rep.asp " & _
"Loaded by " & Session("sp_miLogin")

sql = "sprd_getLastRadicationNumber " & Section
rs.Open sql,cn,3
Inicio = Request.Form.Item("Inicio")
Fin = rs.Fields(0)
rs.Close 

write_sp_log cn, 13291, "sprd_getLastRadicationNumber", 0, "", "", 0, 0, "", "rad_rep.asp " & _
"Loaded by " & Session("sp_miLogin")

rs.Open sql,cn,3
FinOld = rs.Fields(0)
rs.Close 

write_sp_log cn, 13291, "sprd_getLastRadicationNumberOld", 0, "", "", 0, 0, "", "rad_rep.asp " & _
"Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"rad_rep.asp","rad_rep.asp " & "Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"sprd_get_RadUserSection " & Session.Contents("sp_idworker")&" - "&"sprd_getLastRadicationNumber " & Section,"N/A","null","Consulta","N/A"

set rs = Nothing
closeconnpipelinedb
set cn= Nothing

OpenHTML
OpenHead
       
%>
<SCRIPT LANGUAGE=javascript src=../_pipeline_scripts/validation.js></SCRIPT>
</head>
<body class="cuerpo">
    <div class="contenido">
		<br><br>
		<div class="subtituloPagina">
			Reporte Radicación por Número
		</div>
		<br><br>
<%

otbl"tblcontenido"
    opentr""
        opentd"",""
            if Section <> "-1" then
            otbl"tblvalores"   
                    opentr""
                        opentd"",""
                            Response.Write "<br>"
                        closetd
                   closetr             
                OpenForm "Report", "post", "rad_rep_initial.asp", "onSubmit='return formValidation(this)'"
                    opentr ""
                        opentd"'titulotabla'","colspan=2"
                            Response.Write "Reporte de Radicación por Número de Solicitud y Piso"			
                        closetd
                    closetr	
                    OpenTr ""
			            OpenTd "'separadorSecciones'", "colspan=2 align=center"									
			            CloseTd
		            CloseTr
	                OpenTr ""
		                OpenTd "'labels'", "align=left"
			                Response.Write "Desde Solicitud No."
		                CloseTd
		                OpenTd "''", "align=left"
			                PlaceInput "Inicio","text",Inicio,"class=txboxgenericas id='RN          P  No. Solicitud Inicial'"
		                CloseTd
	                CloseTr
	                OpenTr ""
		                OpenTd "'labels'", "align=left"
			                Response.Write "Hasta Solicitud No."
		                CloseTd
		                OpenTd "''", "align=left"
			                Response.Write Fin
			                PlaceInput "Fin","hidden",Fin,""
		                CloseTd
	                CloseTr
                    opentr""
                        opentd"",""
                            Response.Write "<br>"
                        closetd
                    closetr
	                OpenTr ""
		                OpenTd "''", "align=center colspan=2"
				            PlaceInput "Enviar","submit","Enviar","class=button-OLD "
		                CloseTd
	                CloseTr
	                PlaceInput "Seccion","Hidden",Section,""
	                PlaceInput "idSociedad","hidden",Session.Contents("sp_idSoc"),""
                CloseForm
            ctbl

            otbl"tblvalores"
                OpenForm "Report", "post", "rad_rep_initial_1.asp", "onSubmit='return formValidation(this)'"
                    opentr""
                        opentd"",""
                            Response.Write "<br><br>"
                        closetd
                    closetr
                    opentr ""
                        opentd"'titulotabla'","colspan=2"
                            Response.Write "Reporte de Radicación por Número de Solicitud - Sistema anterior (Pipeline 1.0)"			
                        closetd
                    closetr	
                    OpenTr ""
			            OpenTd "'separadorSecciones'", "colspan=2 align=center"									
			            CloseTd
		            CloseTr
	                OpenTr ""
		                OpenTd "'labels'", "align=left"
			                Response.Write "Desde Solicitud No."
		                CloseTd
		                OpenTd "''", "align=left"
			                PlaceInput "Inicio","text",Inicio,"class=txboxgenericas id='RN          P  No. Solicitud Inicial'"
		                CloseTd
	                CloseTr
	                OpenTr ""
		                OpenTd "'labels'", "align=left"
			                Response.Write "Hasta Solicitud No."
		                CloseTd
		                OpenTd "''", "align=left"
			                PlaceInput "Fin","text",FinOld,"class=txboxgenericas id='RN          P  No. Solicitud Final'"
		                CloseTd
	                CloseTr
                    opentr""
                        opentd"",""
                            Response.Write "<br>"
                        closetd
                    closetr
	                OpenTr ""
		                OpenTd "''", "align=center colspan=2"
			                PlaceInput "Enviar","submit","Enviar","class=button-OLD "
		                CloseTd
	                CloseTr
	                PlaceInput "Seccion","Hidden",Section,""
	                PlaceInput "idSociedad","hidden",Session.Contents("sp_idSoc"),""
                CloseForm
            ctbl

            otbl "tblvalores"
                 opentr""
                    opentd"",""
                        Response.Write "<br><br>"
                    closetd
                closetr
                opentr ""
                    opentd"'titulotabla'","colspan=5"
                        Response.Write "Reporte de Radicación por Identificación"			
                    closetd
                closetr	
                OpenTr ""
			        OpenTd "'separadorSecciones'", "colspan=5 align=center"									
			        CloseTd
		        CloseTr
                OpenForm "Report", "post", "rad_rep_initial_id.asp", "onSubmit='return formValidation(this)'"
	                OpenTr ""
		                OpenTd "'labels'", "align=left"
			                Response.Write "Identificación No."
		                CloseTd
		                OpenTd "''", "align=left, "
			                PlaceInput "Identification","text",Inicio,"class=txboxgenericas id='RN          P  No. Solicitud Inicial'"
			                Response.Write DocTypeCombo
		                CloseTd
		                OpenTd "'labels'", "align=left"
			                Response.Write "WorkFlow : "
		                CloseTd
		                OpenTd "''", "align=left"
			                OpenCombo "SistemaWF",  "class=listagenerica"
					                PlaceItem "", 2, "Pipeline 2.0"
					                PlaceItem "", 1, "Pipeline 1.0"
			                CloseCombo
		                CloseTd
	                CloseTr
                    opentr""
                        opentd"",""
                            Response.Write "<br>"
                        closetd
                    closetr
                    OpenTr ""
		                OpenTd "''", "align=center colspan=5"
				                PlaceInput "Enviar","submit","Enviar","class=button-OLD "
		                CloseTd
	                CloseTr
                    PlaceInput "Seccion","Hidden",Section,""
	                PlaceInput "idSociedad","hidden",Session.Contents("sp_idSoc"),""
                CloseForm
            ctbl
        else
            otbl"tblvalores"
                opentr""
                    opentd"","texto-informativo"
                        Response.Write Application.Contents("error4")
                    closetd
                closetr
            ctbl
        end if
        closetd
    closetr
ctbl

%>
        <p/>
        <p/>
		</div>
	</body>
</html>
<%

If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If

%>