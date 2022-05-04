<%@ Language=VBScript %>
<%
'Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../../_ScriptLibrary/pm.asp"-->
<!--#include file="../_pipeline_scripts/url_check.asp"-->
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../../operations/transactions/reglamentoscriptsPipeline.asp"-->
<%
Authorize 1,16


Dim Sql 
Dim adoConn 
Dim rstDescription 
Dim rs, cn, ProdName, ProvFund
Dim arrContract
Dim Contract, Product, Plan, ClientId, DocType, Name, Phone

dim Reference 

Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
Name = Request.Form("Name")
ClientId = Request.Form("ClientId")
Session("name")=ClientId
DocType = Request.Form("DocType")
Phone = Request.Form("Phone")

write_dataLog Response.Status,"contract_info_mfcr.asp","process_selection Loaded by: " & Session("sp_miLogin") & " selected  for the Contract " & Contract,ClientId,"ContratoRetirosPAC_GetForContrato " & Contract & " - ContratoRetirosPAC_GetForContrato " & Contract,"N/A","null","Consulta","N/A"


Session("seleccionContrato")="false"

'Connect to database
Set adoConn = GetConnpipelineDB
Set rstDescription = Server.CreateObject("ADODB.Recordset")


Sql = "spsp_ComplemnetData_GetByClient '" & CStr(DocType) & "', " & CStr(ClientId)
rstDescription.Open Sql, adoConn

If rstDescription.BOF And rstDescription.EOF Then
	arrComplement = 0
Else
	arrComplement = rstDescription.GetRows()
	If IsArray(arrComplement) Then
		IsCoreDescription = arrComplement(0,0)	
	End If
End If
rstDescription.Close
 Reference = GetReferenciaUnica("MFCR", Contract)

'Get Contract description
Sql = "sppl_GetDetallesContrato " & Contract & _
", '" & Product & "', '" & Plan & "'"
rstDescription.Open Sql, adoConn
If rstDescription.BOF And rstDescription.EOF Then
	arrContract = 0
Else
	arrContract = rstDescription.GetRows()
End If
rstDescription.Close

if  IsArray(arrContract) then
	if not (isnull(arrContract(15,0)) or (isnull(arrContract(35,0)))) then
			AuthorizeContractAccess Cstr(Session("sp_AccessLevel")),  CStr(Session("sp_IdAgte")), _
						CStr(arrContract(15,0)), CStr(arrContract(35,0)) 
	end if
end if


write_sp_log adoConn, 12100, "sppl_GetDetallesContrato", Contract, Product, Plan, ClientId, 0, "", "contract_info_other.asp " & _
"- Loaded by" & Session("sp_miLogin")

If IsArray(arrContract) Then
	'Get Product description
	Sql = "spem_GetDescriptionProduct " & arrContract(0,0)
	rstDescription.Open Sql, adoConn
	If rstDescription.BOF And rstDescription.EOF Then
		ProdName = arrContract(0,0)
	Else
		ProdName = rstDescription.Fields("Descripcion")
	End if
	rstDescription.Close

	write_sp_log adoConn, 12100, "spem_GetDescriptionProduct", Contract, Product, Plan, ClientId, 0, "", "contract_info_other.asp " & _
	"- Loaded by" & Session("sp_miLogin")
End If

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	OpenBody "''", "bgcolor='#FFFFFF' text='#000000'"
	If InStr(1, Request.ServerVariables("HTTP_REFERER"), "menu.asp") = 0 Then
		'Reload Left Menu -- START
		OpenForm "menu_left", "post", "../menu/menu.asp", "target=menu"
			PlaceInput "Name", "hidden", Request.Form("Name"), ""
			PlaceInput "ClientId", "hidden", ClientId, ""
			PlaceInput "DocType",  "hidden", DocType, ""
			PlaceInput "Phone", "hidden", Phone, ""
			PlaceInput "Contract", "hidden", Contract, ""
			PlaceInput "Product", "hidden", Product, ""
			PlaceInput "Plan", "hidden", Plan, ""
'			PlaceInput "OU", "hidden", OU, ""
		CloseForm
		
%>
		<script language=javascript>
			document.menu_left.submit();
		</SCRIPT>
<%
End If

OpenTable "90%","'' align=center"
	OpenTr ""
		OpenTd "",""
			OpenTable "100%","'' align=center"
				OpenTr ""
					OpenTd "h3",""
						Response.Write "Información del crédito" 
					CloseTd
				CloseTr
			CloseTable
			OpenTable "100%",""
				OpenTr ""			
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			
				OpenTr ""
					OpenTd "thead","width=30%"
						Response.Write "Número crédito:"
					CloseTd
					OpenTd "tbody", "width=5%"
						Response.Write "&nbsp;"
					CloseTd										
					OpenTd "tbody","width=15%"
						If IsNull(Reference) or len(Reference)>12 Then
							Response.Write Request.QueryString("Contract")
						Else
							response.Write Reference
						End if
					CloseTd
					OpenTd "tbody","width=50%"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Fecha desembolso/renovación:"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
					    If IsNull(arrContract(6,0)) or len(arrContract(6,0))<1 Then
					    Response.Write ""
						Else
						Response.Write FormatDateTime(arrContract(6,0),2)
						End if
												
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Fecha vencimiento:"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
						If IsNull(arrContract(8,0)) or len(arrContract(8,0))<1 Then
					    Response.Write ""
						Else
						Response.Write FormatDateTime(arrContract(8,0),2)
						End if
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				
				'Espacio definido en el diseño
				OpenTr ""
					OpenTd "thead",""
						Response.Write "&nbsp;"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Valor desembolsado:"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
					    If IsNull(arrContract(19,0)) or len(arrContract(19,0))<1 Then
					    Response.Write "$0.0"
						Else
						Response.Write FormatCurrency(arrContract(19,0),2)
						End If
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Saldo Capital:"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
					    If IsNull(arrContract(16,0)) or len(arrContract(16,0))<1 Then
					    Response.Write "$0.0"
						Else
						Response.Write FormatCurrency(arrContract(16,0),2)
						End If
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Tasa (E.A.):"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
						If IsNull(arrContract(105,0)) or len(arrContract(105,0))<1 Then
					    Response.Write "%"
						Else
						Response.Write arrContract(105,0) + "%"
						End If					
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Saldo intereses a la fecha:"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
					if IsNull(arrContract(17,0)) or len(arrContract(17,0))<1 Then
					    Response.Write "$0.0"
						else
						Response.Write FormatCurrency(arrContract(17,0),2)
						end if
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Saldo intereses al vencimiento:"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
					if IsNull(arrContract(20,0)) or len(arrContract(20,0))<1 Then
					    Response.Write "$0.0"
						else
						Response.Write FormatCurrency(arrContract(20,0),2)
						end if
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "thead",""
						Response.Write "Saldo intereses de mora:"
					CloseTd
					OpenTd "tbody", ""
						Response.Write "&nbsp;"
					CloseTd															
					OpenTd "tbody",""
						if IsNull(arrContract(31,0)) or len(arrContract(31,0))<1 Then
					    Response.Write "$0.0"
						else
						Response.Write FormatCurrency(arrContract(31,0),2)
						end if						
					CloseTd
					OpenTd "tbody",""
						Response.Write "&nbsp;"
					CloseTd
				CloseTr
			CloseTable
		CloseTd
	CloseTr
CloseTable

Response.Write "</body>" & vbCrLf & _
"</html>"
Set rstDescription = Nothing
'Close Database connection
CloseConnpipelineDB
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
