<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:					consolidate.asp 8900
'Path:							management/
'Created By:					A. Orozco 2001/09/05
'Last Modified:				A. Orozco 2001/09/19
'									A. Orozco 2001/10/11
'Modifications:				File Creation
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
<!--#include file="../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<%
Authorize 1,12
Dim objConn 'Database Connection
Dim objRst 'Recordset object
Dim strSQL 'Query container
Dim Sel, I
Set objConn = GetConnPipelineDB

write_sp_log objConn, 8900, "", 0, "", "", 0, 0, "", "management/consolidate.asp Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"consolidate.asp", "management/consolidate.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"","N/A","null","Consulta","N/A"

CloseConnPipelineDB
OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src=../operations/_pipeline_scripts/validation.js></SCRIPT>
</head>
<body class="cuerpo">
    <div class="contenido">
	</br></br>
		<div class="subtituloPagina">
			Consolidar
		</div>
		</br></br>
<%
		PlaceLink "REL", "stylesheet", "../css/OLDMutualStyle.css", "text/css"
		
	otbl "'tblContenido'"	
        openTr ""
            openTd "",""	
                otbl "'TablaValores'"	
			        OpenForm "sel_unit", "post", "consolidation_detail.asp", "onSubmit='javascript:return formValidation(this)'"
			        Response.Write "<br><br>"
                    OpenTr ""
				        OpenTd "'labels'", "align=center"
					        Response.Write "Unidad "
                        CloseTd
                        OpenTd "''", "align=left width=50%"
					        OpenCombo "opUnit", "class=listaFecha"
						        For I = 0 To 5
							        If I = 0 Then
								        Sel = "selected"
							        Else
								        Sel = ""
							        End If
							        PlaceItem Sel, "00000" & I, I
						        Next
					        CloseCombo
				        CloseTd
			        CloseTr           
			        OpenTr "align = center"                
				        OpenTd "''", "align=center colspan=2"
                         Response.Write "<br>"
					        PlaceInput "go", "submit", "Continuar", "class=button-OLD"
				        CloseTd
			        CloseTr
			        CloseForm
		        ctbl
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