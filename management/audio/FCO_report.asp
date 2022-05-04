<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:						FCO_report.asp 1600
'Path:							management/audio
'Created By:					Fabio calvache 2002/10/10
'Modifications:					File Creation
'	Andrés Jaramillo				Ajustes Marca 05/02/2014
'Returns:						Transacctions
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
Authorize 7,20
PlaceLink "REL", "stylesheet", "../../css/OLDMutualStyle.css", "text/css"
Dim I, Sel

write_dataLog  Response.Status,"FCO_report.asp", "FCO_report.asp loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"null","N/A","null","Consulta","N/A"

%>
<HTML>
	<head>
		<meta name="" http-equiv="expires" content="Wednesday, 27-Dec-95 05:29:10 GMT"/>
		<meta name="" http-equiv="Pragma" content="no_cache"/>
<SCRIPT LANGUAGE=javascript src='../_pipeline_scripts/validation.js'></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
function chkDate(form) {
	var sDate;
	var qDate;
	sDate = new Date(form.INIANNO.value + form.INIMES.value + form.INIDIA.value);
	qDate = new Date(form.FINANNO.value + form.FINMES.value + form.FINDIA.value);
	if (qDate - sDate < 0) {
		alert('La fecha inicial debe ser menor que la fecha final');
		return false;
	};
	return true;
}
//-->
</SCRIPT>
	</head>
	<body class="cuerpo">
		<div class="Contenido">
			<div class="subtituloPagina">Reporte Transferencias FIC´S</div>

<%
OpenForm "date", "post", "FCO_reportsubmit.asp", "onSubmit='return formValidation(this)' id='date'"
otbl"'tblContenido'"
	opentr""
		opentd"",""
			otbl"'tblValores'"
				OpenTr ""
					OpenTd "'labels'", ""
						Response.Write "Fecha inicial"
					CloseTd
					OpenTd "", ""
						Response.Write "Año: "
						PlaceInput "INIANNO", "text", Year(Now), "class=txboxgenericas size=4 maxlength=4 id='RN           Y Año'"
						Response.Write "Mes: "
						OpenCombo "INIMES", "class=listafecha id='             M Mes'"
							For I = 1 To 12
								If Month(Now) = I Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, I, I
							Next
						CloseCombo
						Response.Write "Día: "
						PlaceInput "INIDIA", "text", Day(Now), "class=txboxgenericas size=2 maxlength=2 id='RN           D Día'"
						Response.Write "<br>"
					CloseTd
				CloseTr
				OpenTr "align=center"
					OpenTd "'labels'", ""
						Response.Write "Fecha final"
					CloseTd
					OpenTd "", ""
						Response.Write "Año: "
						PlaceInput "FINANNO", "text", Year(Now), "class=txboxgenericas size=4 maxlength=4 id='RN           Y Año'"
						Response.Write "Mes: "
						OpenCombo "FINMES", "class=listafecha id='             M Mes'"
							For I = 1 To 12
								If Month(Now) = I Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, I, I
							Next
						CloseCombo
						Response.Write "Día: "
						PlaceInput "FINDIA", "text", Day(Now), "class=txboxgenericas size=2 maxlength=2 id='RN           D Día'"
						Response.Write "<br>"
					CloseTd
				CloseTr
			ctbl
		closetd
	closetr
	OpenTr ""
			Response.Write "<br>"
	CloseTr
	OpenTr ""
		OpenTd "", ""
			PlaceInput "Send", "submit", " Enviar ", "class=button-OLD onClick='javascript:return chkDate(form)'"
		CloseTd
	CloseTr
ctbl
CloseForm
CloseBody
CloseHTML

If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
	</div>

</body>
<HTML>