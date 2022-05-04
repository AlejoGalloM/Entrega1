<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:						calls_report.asp 1600
'Path:							management/audio
'Created By:					Fabio calvache 2002/10/10
'Modified by:
'   Armando J. Arias Gómez			2008/05/07 - PlaceTitle
'	Andrés Jaramillo				Ajustes Marca 05/02/2014
'Modifications:					File Creation
'Returns:						Calls per Hour
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
Dim I, Sel

write_dataLog  Response.Status,"calls_report.asp", "calls_report.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"null","N/A","null","Consulta","N/A"

%>
<HTML>
	<head>
		<meta name="" http-equiv="expires" content="Wednesday, 27-Dec-95 05:29:10 GMT"/>
		<meta name="" http-equiv="Pragma" content="no_cache"/>
		<link href="../../css/style.css" rel="stylesheet" type="text/css"/>
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
		<div class="mainHeader1">
			<div class="title1" style="width: 100%">
				Administración de Audio</div>
		</div>
		<div class="rounded" width="100%">
			<b class="r1"></b>
			<b class="r2"></b>
			<b class="r3"></b>
			<b class="r4"></b>
		</div>
		<div class="data">
		<div class="header1 master1">
			Reporte Reporte de Llamadas
		</div>
<%
OpenForm "date", "post", "calls_reportsubmit.asp", "onSubmit='return formValidation(this)' id='date'"
OpenTable "50% face=Courier", "'' border=0 align=center"
	OpenTr "class=todd align=center"
		OpenTd "", ""
			Response.Write "<b>Fecha inicial</b>"
		CloseTd
		OpenTd "", ""
			Response.Write "Año: "
			PlaceInput "INIANNO", "text", Year(Now), "class=bttntext size=4 maxlength=4 id='RN           Y Año'"
			Response.Write "Mes: "
			OpenCombo "INIMES", "class=bttntext id='             M Mes'"
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
			PlaceInput "INIDIA", "text", Day(Now), "class=bttntext size=2 maxlength=2 id='RN           D Día'"
			Response.Write "<br>"
		CloseTd
	CloseTr
	OpenTr "class=todd align=center"
		OpenTd "", ""
			Response.Write "<b>Fecha final</b>"
		CloseTd
		OpenTd "", ""
			Response.Write "Año: "
			PlaceInput "FINANNO", "text", Year(Now), "class=bttntext size=4 maxlength=4 id='RN           Y Año'"
			Response.Write "Mes: "
			OpenCombo "FINMES", "class=bttntext id='             M Mes'"
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
			PlaceInput "FINDIA", "text", Day(Now), "class=bttntext size=2 maxlength=2 id='RN           D Día'"
			Response.Write "<br>"
		CloseTd
	CloseTr
CloseTable
OpenTable "50% face=Courier", "'' border=0 align=center"
	OpenTr ""
			Response.Write "<br>"
	CloseTr
	OpenTr "class=todd align=center"
		OpenTd "", ""
			PlaceInput "Send", "submit", " Enviar ", "class=sbttn onClick='javascript:return chkDate(form)'"
		CloseTd
	CloseTr
CloseTable
CloseForm

If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
	</div>
	<div class="rounded" width="100%">
		<b class="r4"></b>
		<b class="r3"></b>
		<b class="r2"></b>
		<b class="r1"></b>
	</div>
</body>
<HTML>
