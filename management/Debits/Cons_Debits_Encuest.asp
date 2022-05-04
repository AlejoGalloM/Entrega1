<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:						Cons_Debits_Encuest 15400
'Path:							management/debits/
'Created By:					R. Lagos 2002/10/16
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
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<%
Authorize 2,20
Dim objConn, cn 'Database Connection
Dim objRst, rs 'Recordset object
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

write_sp_log objConn, 15400, "", Contract, Product, Plan, ClientId, 0, "", "Cons_Debits_Encuest.asp loaded by " & _
Session.Contents("sp_milogin")

write_dataLog Response.Status,"Cons_Debits_Encuest.asp", "Cons_Debits_Encuest.asp Loaded by " & Session("sp_miLogin"),Name,"","N/A","null","Consulta","N/A"

CloseConnPipelineDB
OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src=../../operations/_pipeline_scripts/validation.js></SCRIPT>
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
		</br></br>
		<div class="subtituloPagina">
			</br></br>
			Encuesta Retiros
		</div>
<%
		PlaceLink "REL", "stylesheet", "../../css/OLDMutualStyle.css", "text/css"
	

	otbl "'tblContenido'"	
        openTr ""
            openTd "",""	
                otbl "'TblValores'"			
			        OpenForm "sel_unit", "post", "Debits_Encuest.asp", "onSubmit='javascript:return validate(this)'"
			        OpenTr ""
                        Response.Write "<br><br><br>"
				        OpenTd "'labels'", "align=right valign=middle "
					        Response.Write "Fecha Inicial"
				        CloseTd
				        OpenTd "labelcombo", "align=right "
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
				        OpenTd "''", "'' align=center colspan=8"
                            Response.Write "<br><br><br>"
					        PlaceInput "go", "submit", "Continuar", "class=button-OLD"
				        CloseTd
			        CloseTr
			     CloseForm
		     Ctbl
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