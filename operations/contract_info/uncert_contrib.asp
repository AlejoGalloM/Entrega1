<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		uncert_contrib.asp
'Path:				contract_info/
'Created By:		A. Orozco 2001/07/30
'Last Modified:	A. Orozco 2001/10/03
'						A. Orozco 2001/10/09
'						Guillermo Aristizabal 2001/10/11
'						Rafael lagos 2003/06/10 remove volver button
'Parameters:		User must be logged on
'						Contract number
'Returns:			Gets uncertified contributions from AS400
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
Dim objConn
Dim ObjCon, Pagina, Param, Ctr
Dim Contract, Product, Plan

Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
Ctr = Contract

Authorize 3,1

Set objConn = GetConnPipelineDB

write_sp_log objConn, 300, "", Ctr, Product, Plan, 0, 0, "", "uncert_contrib.asp started loading " & _
"- " & Session("sp_miLogin")

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src='../_pipeline_scripts/validation.js'></SCRIPT>
<%
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
'=======================================================================================
'Start of AS400
'=======================================================================================


If Ctr = "" Then
	Response.Write("Debe existir N�mero de Contrato.")
	Response.End 
End If
	

	''SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNION� EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN
	''Set ObjCon = Server.CreateObject("Conexion.conectar") 
	''FIN SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNION� EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN

'========================================================================================
'Informacion que debe aparecer cuando salga TAX
'========================================================================================
'Param = "http://CODBS02/AS4002/uncert_contrib.asp?Ctr=" & Ctr
'Param = "http://10.42.1.77/400_files/uncert_contrib.asp?Ctr=" & Ctr
'Response.Write "SecondaryContract"
'Response.Write Session.Contents("SecondaryContract")
Product = Trim(Product)
if Trim(Product) = "OMPEV" then
	Param = Application("URLTax") & "?contract=" & Session.Contents("SecondaryContract") & "&function=aportes&product=" & Product 
else 
	Param = Application("URLTax") & "?contract=" & Contract & "&function=aportes&product=" & Product 
end if
'Param = Application("URLTax") & "contract=" & Contract & "&function=aportes&product=" & Product

	''SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNION� EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN
	''Pagina = ObjCon.urlReader(Param)
	''FIN SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNION� EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN
Response.Write "<p>&nbsp;</p><p>&nbsp;</p>"
OpenTable "90%", "'' align=center"
	OpenTr ""
		OpenTd "", ""
'			Response.Write "<p align=center>Consulta de Tax Benefits.</p>"
			'Response.Write Param
			
			''SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNION� EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN
			'response.write Param
			'response.end 
			Set ObjCon = CreateObject("MSXML2.XMLHTTP")
			ObjCon.Open "GET", Param , false
			ObjCon.setRequestHeader "Content-Type","text/xml"
			ObjCon.Send(null)
			Pagina = ObjCon.responseText
			Response.Write Pagina
			''FIN SE MODIFICA DADO QUE EL CONEXION.CONECTAR NO FUNION� EN EL NUEVO SERVIDOR 20201202 ALEJANDRO PULGARIN
'========================================================================================
'Informacion que debe aparecer cuando salga TAX
'========================================================================================
		CloseTd
	CloseTr
'=======================================================================================
'End of AS400
'=======================================================================================
	OpenTr ""
		OpenTd "''", "align=center"
			OpenForm "","Post","contract_info_mfund.asp", "onSubmit=formValidation(this)"
				PlaceInput "Contract", "hidden", Contract, ""
				PlaceInput "Product", "hidden", Product, ""
				PlaceInput "Plan", "hidden", Plan, ""
				'PlaceInput "Back", "Submit","Volver","class=sbttn"
			CloseForm
		CloseTd
	CloseTr
CloseTable

write_sp_log objConn, 300, "", Ctr, Product, Plan, 0, 0, "", "uncert_contrib.asp finished loading " & _
"- " & Session("sp_miLogin")

write_dataLog Response.Status,"uncert_contrib.asp","uncert_contrib.asp finished loading - " & Session("sp_miLogin"),Session.contents("name"),pagina,"N/A","null","Consulta","N/A"

CloseConnPipelineDB
Set objConn = Nothing
Response.Write "<p>&nbsp;</p>" & vbCrLf
CloseBody
CloseHTML
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
