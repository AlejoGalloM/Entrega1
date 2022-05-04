<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		contract_info_proto.asp 12100
'Path:				contract_info/
'Created By:		J Moreno 2003/10/30
'Last Modified:		
'Parameters:		User must be logged on
'						Session("docNumber")
'Returns:			contract information for protipo products 
'Additional Information: redirect to .net
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
Authorize 1,16

Dim Sql 'SQL Sentences holder
Dim adoConn 'Database Connection
Dim rstDescription 'Results recordset
Dim rs, cn, ProdName, ProvFund
Dim arrContract, arrHIstory
Dim arrAsset
Dim arrStanding
Dim Contract, Product, Plan, ClientId, DocType, Name, Phone, DateAfiliation, FP, DateEnter, saldo
Dim I,J 'Asset & Standing Alloc counters
Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
Name = Request.Form("Name")
ClientId = Request.Form("ClientId")
Session("name")=ClientId
DocType = Request.Form("DocType")
Phone = Request.Form("Phone")

write_dataLog Response.Status,"contract_info_proto.asp","process_selection Loaded by: " & Session("sp_miLogin") & " selected  for the Contract " & Contract,ClientId,"ContratoRetirosPAC_GetForContrato " & Contract & " - ContratoRetirosPAC_GetForContrato " & Contract,"N/A","null","Consulta","N/A"


Session("seleccionContrato")="false"
'Connect to database
Set adoConn = GetConnpipelineDB
Set rstDescription = Server.CreateObject("ADODB.Recordset")

'response.end

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

'response.write sql



if  IsArray(arrContract) then
	if not (isnull(arrContract(15,0)) or (isnull(arrContract(35,0)))) then
			AuthorizeContractAccess Cstr(Session("sp_AccessLevel")),  CStr(Session("sp_IdAgte")), _
						CStr(arrContract(15,0)), CStr(arrContract(35,0)) 
	end if
end if


'write_sp_log(connection, page_id, sp, contract, product, plan, client_id, error, conf_num, text)
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
	'write_sp_log(connection, page_id, sp, contract, product, plan, client_id, error, conf_num, text)
	write_sp_log adoConn, 12100, "spem_GetDescriptionProduct", Contract, Product, Plan, ClientId, 0, "", "contract_info_other.asp " & _
	"- Loaded by" & Session("sp_miLogin")
End If



Set rstDescription = Nothing
'Close Database connetction
CloseConnpipelineDB
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
		'Reload Left Menu -- END
	End If
CloseBody
CloseHTML


'Reload Left Menu -- START
OpenForm "menu", "post", "../menu/menu.asp", "target=menu"
	PlaceInput "Name", "hidden", Name, ""
	PlaceInput "ClientId", "hidden", ClientId, ""
	PlaceInput "DocType",  "hidden", DocType, ""
	PlaceInput "Contract", "hidden", Contract, ""
	PlaceInput "Product", "hidden", Product, ""
	PlaceInput "Phone", "hidden", Phone, ""
	PlaceInput "Plan", "hidden", Plan, ""
	PlaceInput "Option", "hidden", 0, ""
CloseForm
%>
<script language=javascript>
	document.menu.submit();
</SCRIPT>
<%
'Reload Left Menu -- END
OpenForm "contract", "post", Application("URLProto"), ""
	PlaceInput "Name", "hidden", Name, ""
	PlaceInput "ClientId", "hidden", ClientId, ""
	PlaceInput "DocType",  "hidden", DocType, ""
	PlaceInput "Contract", "hidden", Contract, ""
	PlaceInput "Product", "hidden", Product, ""
	PlaceInput "Phone", "hidden", Phone, ""
	PlaceInput "Plan", "hidden", Plan, ""
	PlaceInput "Option", "hidden", 0, ""
	if  IsArray(arrContract) then
		if not (isnull(arrContract(11,0))) then
			PlaceInput "DateAfiliation", "hidden", FormatDateTime(arrContract(11,0),2), ""
		else
			PlaceInput "DateAfiliation", "hidden", "N/D", ""
		end if
		
		if (isnull(arrContract(36,0))) then
			PlaceInput "FP", "hidden", arrContract(36,0), ""
		else
			PlaceInput "FP", "hidden", "N/D", ""
		end if
		
		if (isnull(arrContract(6,0))) then
			PlaceInput "DateEnter", "hidden", FormatDateTime(arrContract(6,0),2), ""
		else
			PlaceInput "DateEnter", "hidden", "N/D", ""
		end if
		
		if (isnull(arrContract(9,0))) then
			PlaceInput "Saldo", "hidden", FormatDateTime(arrContract(9,0),2), ""
		else
			PlaceInput "Saldo", "hidden", "N/D", ""
		end if
	else
	PlaceInput "DateAfiliation", "hidden", "N/D", ""
	PlaceInput "FP", "hidden", "N/D", ""
	PlaceInput "DateEnter", "hidden", "N/D", ""
	PlaceInput "Saldo", "hidden", "N/D", ""
	end If
CloseForm



'response.write Application("URLProto")
'response.write Err.number0
'response.end


%>
<script language=javascript>
	document.contract.submit();
</SCRIPT>
<%


If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
