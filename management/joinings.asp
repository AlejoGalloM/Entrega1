<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:					joinings.asp 8800
'Path:							management/
'Created By:					A. Orozco 2001/09/13
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
Authorize 0,12
PlaceLink "REL", "stylesheet", "../css/OLDMutualStyle.css", "text/css"
Dim objConn 'Database Connection
Dim objRst 'Recordset object
Dim strSQL 'Query container
Dim Sel, I
Dim Contract, Product, Plan, ClientId, Name

Contract = Request.Form("Contract")
If Contract = "" Then Contract = 0
Product = Request.Form("Product")
Plan = Request.Form("Plan")
ClientId = Request.Form("ClientId")
If ClientId = "" Then ClientId = 0
Name = Request.Form("Name")
Set objConn = GetConnPipelineDB

write_sp_log objConn, 8800, "", Contract, Product, Plan, ClientId, 0, "", "joinings.asp loaded by " & _
Session.Contents("sp_milogin")

write_dataLog Response.Status,"joinings.asp","joinings.asp" & " Loaded by " & Session("sp_miLogin"),Name,"","N/A","null","Consulta","N/A"


CloseConnPipelineDB
OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src=../operations/_pipeline_scripts/validation.js></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
	function validate(form) {
		var value = false
		with (form) {
			sendForm(form)
			value = dateValidation(s_yval.value, s_mval.value, s_dval.value);
			value = value && dateValidation(e_yval.value, e_mval.value, e_dval.value);
		}
		if (!value) {
			enableButtons(form)
		}
		return value
	}
//-->
</SCRIPT>
</head>
<body class="cuerpo">
    <div class="contenido">
		<div class="subtituloPagina">
			Afiliaciones
		</div>

<%
	otbl "'tblContenido'"	
        openTr ""
            openTd "",""	
                otbl "'TablaValores'"	
		            OpenForm "sel_unit", "post", "joinings_detail.asp", "onSubmit='javascript:return validate(this)'"
                    Response.Write "<br><br>"
			            OpenTr ""
				            OpenTd "'labels'", "align=right valign=middle "
					            Response.Write "Fecha Inicial"
				            CloseTd
				            OpenTd "'labelcombo'", "align=right "
					            Response.Write "Año: "
					            OpenCombo "s_yval", "class=listaFecha"
						            For I = 1990 To 2030
							            If Year(Date) = I Then
								            Sel = "selected"
							            Else
								            Sel = ""
							            End If
							            PlaceItem sel, I, I
						            Next
					            CloseCombo
					            Response.Write "Mes: "
					            OpenCombo "s_mval", "class=listaFecha"
						            For I = 1 To 12
							            If Month(Date) = I Then
								            Sel = "selected"
							            Else
								            Sel = ""
							            End If
							            PlaceItem sel, MonthName(I, 1), I
						            Next
					            CloseCombo
					            Response.Write "Día: "
					            OpenCombo "s_dval", "class=listaFecha"
						            For I = 1 To 31
							            If Day(Date) = I Then
								            Sel = "selected"
							            Else
								            Sel = ""
							            End If
							            PlaceItem sel, I, I
						            Next
					            CloseCombo
				            CloseTd
			            CloseTr
			            OpenTr ""
				            OpenTd "'labels'", "align=right valign=middle"
					            Response.Write "Fecha Final"
				            CloseTd
				            OpenTd "'labelcombo'", "align=left"
					            Response.Write "Año: "
					            OpenCombo "e_yval", "class=listaFecha"
						            For I = 1990 To 2030
							            If Year(Date) = I Then
								            Sel = "selected"
							            Else
								            Sel = ""
							            End If
							            PlaceItem sel, I, I
						            Next
					            CloseCombo
					            Response.Write "Mes: "
					            OpenCombo "e_mval", "class=listaFecha"
						            For I = 1 To 12
							            If Month(Date) = I Then
								            Sel = "selected"
							            Else
								            Sel = ""
							            End If
							            PlaceItem sel, MonthName(I, 1), I
						            Next
					            CloseCombo
					            Response.Write "Día: "
					            OpenCombo "e_dval", "class=listaFecha"
						            For I = 1 To 31
							            If Day(Date) = I Then
								            Sel = "selected"
							            Else
								            Sel = ""
							            End If
							            PlaceItem sel, I, I
						            Next
					            CloseCombo
				            CloseTd
			            CloseTr
                        OpenTr ""
				            OpenTd "'labels'", "align=right"
					        closeTd
							OpenTd "'labelcombo'", "align=center"
								Response.Write "<b>Unidad : </b>"
							    OpenCombo "opUnit", "class=listaFecha"
						            For I = 0 To 5
							            If I = 0 Then
								            Sel = "selected"
							            Else
								            Sel = ""
							            End If
							            If I = 0 Then
								            PlaceItem Sel, I, "Cualquiera"
							            Else
								            PlaceItem Sel, I, I
							            End If
						            Next
					            CloseCombo
				            CloseTd
			            CloseTr
			            OpenTr ""
				            OpenTd "''", "colspan=2 align=center"
                                Response.Write "<br><br>"
					            PlaceInput "go", "submit", "Continuar", "class=button-OLD"
				            CloseTd
			            CloseTr
			        CloseForm
		        Ctbl
                closeTd
        CloseTr
        Ctbl
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