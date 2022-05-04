<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:						Mfund_report.asp 1600
'Path:							management/audio
'Created By:					Rafael Lagos 2003/01/15
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
Dim objConn 'Database Connection
Dim objRst,rs 'Recordset object
Dim strSQL 'Query container
Dim Sel, I, cn
Dim Contract, Product, Plan, ClientId, Name

Contract = Request.Form("Contract")
If Contract = "" Then Contract = 0
Product = Request.Form("Product")
Plan = Request.Form("Plan")
ClientId = Request.Form("ClientId")
If ClientId = "" Then ClientId = 0
Name = Request.Form("Name")
Set objConn = GetConnPipelineDB

write_sp_log objConn, 15400, "", Contract, Product, Plan, ClientId, 0, "", "Cons_Debits_Otros.asp loaded by " & _
Session.Contents("sp_milogin")

write_dataLog  Response.Status,"Mfund_report.asp", "Cons_Debits_Otros.asp loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"","N/A","null","Consulta","N/A"

CloseConnPipelineDB
%>
<HTML>
	<head>
		<meta name="" http-equiv="expires" content="Wednesday, 27-Dec-95 05:29:10 GMT"/>
		<meta name="" http-equiv="Pragma" content="no_cache"/>
		<link href="../../css/style.css" rel="stylesheet" type="text/css"/>
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
			Reporte Transferencias Multifund
		</div>
<%
		OpenTable "90%", "'' align=center border=0"
			OpenForm "sel_unit", "post", "Mfund_reportsubmit.asp", "onSubmit='javascript:return validate(this)'"
'			OpenForm "sel_unit", "post", "Debits_Units_Otros.asp", "onSubmit='javascript:return validate(this)'"
'			OpenTr ""
'				OpenTd "tbody", "align=center colspan=2"
'					Response.Write "Unidad "
'					OpenCombo "opUnit", "class=bttntext"
'						For I = 0 To 5
'							If I = 0 Then
'								Sel = "selected"
'							Else
'								Sel = ""
'							End If
'							If I = 0 Then
'								PlaceItem Sel, I, "Cualquiera"
'							Else
'								PlaceItem Sel, I, I
'							End If
'						Next
'					CloseCombo
'				CloseTd
'			CloseTr
			OpenTr ""
				OpenTd "tbody", "align=center colspan=2"
					Response.Write "&nbsp;"
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "thead", "align=right width=40%"
					Response.Write "Fecha Inicial"
				CloseTd
				OpenTd "tbody", "align=left width=60%"
					Response.Write "Año: "
					OpenCombo "s_yval", "class=bttntext"
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
					OpenCombo "s_mval", "class=bttntext"
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
					OpenCombo "s_dval", "class=bttntext"
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
			OpenTr "valign=middle"
				OpenTd "thead", "align=right width=40% valign=middle"
					Response.Write "Fecha Final"
				CloseTd
				OpenTd "tbody", "align=left width=60%"
					Response.Write "Año: "
					OpenCombo "e_yval", "class=bttntext"
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
					OpenCombo "e_mval", "class=bttntext"
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
					OpenCombo "e_dval", "class=bttntext"
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
				OpenTd "",""
					Response.Write "&nbsp;"
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "tbody", "align=center colspan=2"
					PlaceInput "go", "submit", "Continuar", "class=sbttn"
				CloseTd
			CloseTr
			CloseForm
		CloseTable

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