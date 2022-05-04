<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		units_history.asp 1700
'Path:				history/
'Created By:		A. Orozco 2001/08/21
'Last Modified:	A. Orozco 2001/09/14
'						A. Orozco 2001/10/08
'						Guillermo Aristizabal 2001/10/11
'Parameters:		User must be logged on
'						Contract, Product
'Returns:			Units history for the selected contract-product and specified month-year
'Additional Information:
'===================================================================================
Option Explicit
On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../../_ScriptLibrary/pm.asp"-->
<!--#include file="../_pipeline_scripts/url_check.asp"-->
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<%

Authorize 1,3
Dim rs, cn 'Recordset, Connection
Dim rsFunds
Dim Sql
Dim adoConn
Dim I,J 'Recordset counters
Dim arrFunds, Sel
Dim Contract, Product, Plan
Dim EffDate
Contract =  Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")

Set adoConn = GetConnPipelineDb


Set rsFunds = Server.CreateObject("ADODB.Recordset")
'Get Funds
Sql = "sppl_GetFondosQuery '" + ltrim(Product) +"'"

rsFunds.Open Sql, adoConn, 3
arrFunds = rsFunds.GetRows
rsFunds.Close
Set rsFunds = Nothing

write_sp_log adoConn, 1700, "sppl_GetFondosQuery", Contract, Product, Plan, 0, 0, "", "units_history.asp " & _
"- " & Session("sp_miLogin")

write_dataLog Response.Status,"units_history.asp","units_history.asp " &"- " & Session("sp_miLogin"),Session.contents("name"),"sppl_GetFondosQuery" ,"N/A","null","Consulta","N/A"

CloseConnPipelineDb
Set adoConn = Nothing
OpenHTML
OpenHead
PlaceTitle "Historia Transacciones"
PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src='../_pipeline_scripts/validation.js'></SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
function checkDates(form) {
	sendForm(form);
	with(form) {
		if (!dateValidation(startYear.value, startMonth.value, startDay.value)) {
			enableButtons(form);
			return false
		}
		if (!dateValidation(endYear.value, endMonth.value, endDay.value)) {
			enableButtons(form);
			return false
		}
		return true;
	}
}
//-->
</SCRIPT>
<%
PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
CloseHead
OpenBody "''", "bgcolor='#FFFFFF' text='#000000'"
OpenTable "90%", "'' align=center"
	OpenTr ""
		OpenTd "", ""
			OpenTable "100%", "'' align=center"
				OpenTr ""
					OpenTd "thead", "align=center"
						Response.Write "Valores De Unidad"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead", "align=center"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			CloseTable
			OpenForm "units", "post", "units_history_results.asp", "onSubmit='javascript:return checkDates(this)'"
				PlaceInput "Contract", "hidden", Contract, ""
				PlaceInput "Product", "hidden", Product, ""
				PlaceInput "Plan", "hidden", Plan, ""
			OpenTable "100%", "'' border=0 align=center"
				OpenTr ""
					OpenTd "thead", "width=27% align=center"
						Response.Write "Fecha Inicial"
					CloseTd
					OpenTd "thead", "width=27% align=center"
						Response.Write "Año: "
						OpenCombo "startYear", "class=bttntext"
							For I = 1990 To Year(Date)
								If I = Year(Date) Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, I, I
							Next
						CloseCombo
						Response.Write "Mes: "
						OpenCombo "startMonth", "class=bttntext"
							For I = 1 To 12
								If I = Month(Date) Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, MonthName(I,1), I
							Next
						CloseCombo
						Response.Write "Día: "
						OpenCombo "startDay", "class=bttntext"
							For I = 1 To 31
								If I = Day(Date) Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, I, I
							Next
						CloseCombo
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead", "width=27% align=center"
						Response.Write "Fecha Final"
					CloseTd
					OpenTd "thead", "width=22% align=center"
						Response.Write "Año: "
						OpenCombo "endYear", "class=bttntext"
							For I = 1990 To Year(Date)
								If I = Year(Date) Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, I, I
							Next
						CloseCombo
						Response.Write "Mes: "
						OpenCombo "endMonth", "class=bttntext"
							For I = 1 To 12
								If I = Month(Date) Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, MonthName(I, 1), I
							Next
						CloseCombo
						Response.Write "Día: "
						OpenCombo "endDay", "class=bttntext"
							For I = 1 To 31
								If I = Day(Date) Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, I, I
							Next
						CloseCombo
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead", "width=27% align=center"
						Response.Write "Fondo:"
					CloseTd
					OpenTd "thead", "width=22% align=center"
						OpenCombo "Fund", "class=bttntext"
							For I = 0 To UBound(arrFunds, 2)
								If I = 0 Then
									Sel = "selected"
								Else
									Sel = ""
								End If
								PlaceItem Sel, arrFunds(0, I), arrFunds(1, I)
							Next
						CloseCombo
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "''", "align=right colspan=2"
						PlaceInput "Send", "submit", "Consultar", "class=sbttn"
						PlaceInput "Clear", "reset", "Limpiar", "class=sbttn"
					CloseTd
				CloseTr
			CloseTable
			CloseForm
		CloseTd
	CloseTr
CloseTable
Response.Write "<p>&nbsp;</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf
CloseBody
CloseHTML
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>