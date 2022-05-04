<%@ Language=VBScript %>
<%
'===================================================================================
'@author name:		 		J carreno 
'@exception name:				
'@param name description:	
'@return					results of search for page 12903
'@since						2002/05/22
'@version					1.0
'@File Name:				searchResultsH.asp [12904]
'@Path:						insurance/insurance
'@revision					Julio 22 2002
'@Modified					J Carreño, cajulio@skandia.com.co, add document type 2003/09/05
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
<!--#include file="../../operations/_pipeline_scripts/url_check.asp"-->
<%

if Request.Form("desde") = "EL" then
	Authorize 5,17
else
	Authorize 1,17
end if

Dim strSql 'Stored procedures and SQL queries
Dim objConn 'ADODB Connection
Dim objRst 'ADODB Recordset
Dim Cliente
Dim arrContracts
Dim Contrato
dim plan
dim producto
Dim I, J, K, L 'Used to navigate arrays
Dim QString 'Used to send the form data as a querystring when there are no results
Dim Total 'Display total of records
Dim Name, LastName, Socs, arrSocs, Flag, TotalRadioBtns
dim hacia 

contrato=Request.Form("contract")
cliente=Request.Form("clientid")
producto=Request.Form("product")
plan=Request.Form("plan")

strSql = "Insurance..sp_insu_SearchContractH " & contrato & ", '" & producto & "', '" & plan & "'"
'Response.Write strsql
'Response.End
Set objConn = GetConnpipelineDB
write_sp_log objConn, 13302, "", contrato, producto, plan, Cliente, 0, "", "SP_LOG - Start Search " & Session("sp_miLogin")
write_dataLog Response.Status,"searchresultsh.asp","SP_LOG - Start Search " & Session("sp_miLogin"),Session.contents("name"),"Insurance..sp_insu_SearchContractH " & contrato & "-'" & producto & "'-'" & plan & "'","N/A","null","Consulta","N/A"

Set objRst = Server.CreateObject("ADODB.Recordset")
objRst.Open strSql, objConn
If objRst.BOF And objRst.EOF Then
	arrContracts = 0
	write_sp_log objConn, 13302, "", Contrato, "", "", Cliente, 0, "", "SP_LOG - No results " & Session("sp_miLogin")
Else
	arrContracts = objRst.GetRows()
	write_sp_log objConn, 13302, "", Contrato, "", "", Cliente, 0, "", "SP_LOG - Results Found: " & objRst.RecordCount & " " & Session("sp_miLogin")
End If
objRst.Close

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
		PlaceLink	"REL", "stylesheet", "../../css/style.css", "text/css"
		PlaceLink	"REL", "stylesheet", "../../css/style1.css", "text/css"
		PlaceLink	"REL", "stylesheet", "../../css/estilo.css", "text/css"
	CloseHead
CloseBody

hacia="infoHistoryInsurance.asp"

CloseConnpipelineDB

If IsArray(arrContracts) Then 'Build Table With Data
OpenForm "transaction", "post", hacia, ""

	PlaceInput "desde", "hidden", Request.Form("desde"), ""

	PlaceInput "operation", "hidden", "", ""
	PlaceInput "selection", "hidden", "", ""
   
    placeinput "opcion","hidden",Request.Form("opcion"),""

	Total = UBound(arrContracts, 2) + 1
	OpenTable "100%", "'' border=0"
		OpenTr "class=thead"
			OpenTd "''", ""
				Response.Write "Total: " & Total & " registros"
			CloseTd
			OpenTd "''", "align=right"
				PlaceInput "select", "submit", "Seleccionar", "class=sbttn"
			CloseTd
	CloseTr
	closetable
	OpenTable "100%", "'' border=0"		
		OpenTr "class=thead align=center"
			OpenTd "", " "
				Response.Write "Sel"
			CloseTd
			OpenTd "", " "
				Response.Write "Fecha Inicio"
			CloseTd
			OpenTd "", ""
				Response.Write "Producto/Plan"
			CloseTd
			OpenTd "", ""
				Response.Write "Contrato"
			CloseTd
			OpenTd "", ""
				Response.Write "Identificación"
			CloseTd
			OpenTd "", ""
				Response.Write "Tipo"
			CloseTd
			OpenTd "", ""
				Response.Write "Cliente"
			CloseTd
'=============================================================================================
'=============================================================================================
		CloseTr
	TotalRadioBtns = 0
	For J = 0 To UBound(arrContracts, 2)
		If (J Mod 2) = 0 Then
			OpenTr "class=teven align=center"
		Else
			OpenTr "class=todd align=center"
		End If
	
'=======================================================================================		
			OpenTd "", ""
							PlaceInput "Contract_" & J, "hidden", arrContracts(7,J), ""
							PlaceInput "ClientId_" & J, "hidden", arrContracts(2,J), ""
							PlaceInput "DocType_" & J, "hidden", arrContracts(9,J), ""
							PlaceInput "Product_" & J, "hidden", arrContracts(5,J), ""
							PlaceInput "Plan_" & J, "hidden", arrContracts(6,J), ""
							PlaceInput "Begin_" & J, "hidden", arrContracts(8,J), ""
							Name = arrContracts(0,J) & " " & arrContracts(1,J)
							PlaceInput "Name_" & J, "hidden", Name, ""
							If J = 0 Then
								PlaceInput "Number", "radio", J, "checked id='          R   ' onClick='javascript:document.transaction.submit()'"
							Else
								PlaceInput "Number", "radio", J, " id='          R   ' onClick='javascript:document.transaction.submit()'"
							End If
							TotalRadioBtns = TotalRadioBtns + 1
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(8,J)) Or arrContracts(8,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write  arrContracts(8,J) 
				End If
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(5,J)) Or arrContracts(5,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write  arrContracts(5,J) & "<br>" & arrContracts(6,J)
				End If
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(7,J)) Or arrContracts(7,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write  "<b>" & arrContracts(7,J)& "</b>" & "<br>" & arrContracts(3,J)
				End If
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(2,J)) Or arrContracts(2,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write arrContracts(2,J)
				End If
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(9,J)) Or arrContracts(9,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write arrContracts(9,J)
				End If
			CloseTd
			OpenTd "", ""
				If IsNull(arrContracts(0,J)) Or IsNull(arrContracts(1,J)) Or arrContracts(0,J) = "" Or  arrContracts(1,J) = "" Then
					Response.Write "N/A"
				Else
					Response.Write arrContracts(0,J) & "<br>" & arrContracts(1,J)
				End If
			CloseTd
'=============================================================================================
'=============================================================================================
		CloseTr
	Next
	OpenTr ""
		OpenTd "''", "colspan=10 align=center"
			Response.Write "&nbsp;"
		CloseTd
	CloseTr
	OpenTr ""
		OpenTd "''", "colspan=8 align=center"
			PlaceInput "select", "submit", "Seleccionar", "class=sbttn"
		CloseTd
	CloseTr
	CloseTable
	CloseForm
	If TotalRadioBtns = 0 Then 'Disable Selection buttons
		Response.Write "<script language=javascript>" & vbCrLf & _
		"<!--" & vbCrLf & _
		"	sendForm(document.transaction)" & vbCrLf & _
		"//--></script>"
	End If
	'If there's only one result, go directly to contract_info
	If UBound(arrContracts, 2) = 0 And TotalRadioBtns > 0 Then
%>
		<script language=javascript>
			//document.transaction.submit();
		</SCRIPT>
<%
	End If
Else
	
	'CloseConnpipelineDB
	
	If Err.number <> 0 Then
		Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
		"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
		Server.URLEncode(Err.source))
	End If
	OpenForm "SDN", "post", "searchHistory.asp", ""
	    Response.Write "<br><br>"
		OpenTable "100%", "'' border=0 "
			OpenTr "class=thead"
				OpenTd "''", " align='center'"
					Response.Write "No existen registros hist&oacute;ricos para este contrato"
				CloseTd
			CloseTr
			if Request.Form("desde") = "EL" then
				OpenTr "class=thead"
					OpenTd "''", " align='center'"
						placeinput "button1","button","Regresar"," class='sbttn2' onclick=window.location='../insurance/default.asp'"
					CloseTd
				CloseTr				
			end if			
		closetable
	CloseForm
End If
	CloseBody
CloseHTML



If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>