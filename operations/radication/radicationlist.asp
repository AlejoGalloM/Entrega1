<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:					radication.asp 1100
'Path:							/operations/radication/radicationlist.asp
'Created By:					Guillermo Pinerez 2001/08/22
'Last Modified:				A. Orozco 2001/09/18
'									A. Orozco 2001/10/10
'				Guillermo Aristizabal 2001/10/11
'Modifications:				
'Parameters:						Contract
'									Unit
'									ClientId
'									Product
'									AgentId
'Returns:						
'Additional Information:	
'===================================================================================
Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../../operations/transactions/reglamentoscriptsPipeline.asp"-->
<!--#include file="radicationqueries.asp"-->
<%
Authorize 3,2
Response.Write "<link rel='stylesheet' href='../../css/style.css' type='text/css'>" & vbCrLf
'== declares ===
Dim Product
Dim Contract
Dim Plan
dim nroDocum
dim DocType
dim idType
dim Name
dim ProcessName
dim arrRadication
dim arrRadicationDetails
dim LastRadID
dim FirstRadID
dim RadDetails
dim UpDown ' details are goin' down or up
dim J,I
dim classid
Dim objConn ' I&T - DMPC
Dim Reference ' I&T - DMPC
dim objRst



'== initials asignments ==
Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
nroDocum = Request.Form("ClientId")
DocType = Request.Form("DocType")
LastRadID = Request.Form("LastRadID")
FirstRadID = Request.Form("FirstRadID")
RadDetails = Request.Form("RadDetails")
UpDown = Request.Form("UpDown")



Set objConn = GetConnPipelineDB

write_sp_log objConn, 1100, "", Contract, Product, Plan, 0, 0, "", "radicationlist.asp loaded " & _
"by " & Session("sp_miLogin")

write_dataLog Response.Status,"radicationlist.asp","radicationlist.asp " &"- " & Session("sp_miLogin"),Session.contents("name"),"null" ,"N/A","null","Consulta","N/A"


CloseConnPipelineDB
set objRst= Server.CreateObject("ADODB.Recordset")
set objConn=GetConnpipelineDB



if len(UpDown) = 0 then
	arrRadication = getRadicationList(nroDocum, DocType , 0, "Down" )
else
	if UpDown = "Up" then
		arrRadication = getRadicationList(nroDocum, DocType ,  FirstRadID, UpDown  )
	else
		arrRadication = getRadicationList(nroDocum, DocType , LastRadID, UpDown  )
	end if
end if

'response.end

if IsArray( arrRadication ) then
	if len(RadDetails) <> 0 then
		arrRadicationDetails = RadicationDetails( RadDetails )
	else
		arrRadicationDetails = RadicationDetails( arrRadication(0,0) )
	end if
else
	Response.Write "<p><p><p>"
	OpenTable "70%", "'' align=center border=0"
		OpenTr "thead"
			OpenTd "teven", "align=left "
				Response.Write "<p><p><p><center><strong>No ha radicado nada</center>"
			CloseTd
		CloseTr
	CloseTable

end if




OpenHTML
OpenBody "",""
	OpenTable "70%", "'' align=center border=0"
		OpenTr "todd"
			OpenTd "thead", "align=left "
				Response.Write "Proceso"
			CloseTd
			OpenTd "thead", "align=left width=500px"
				Response.Write arrRadicationDetails(10,0)
			CloseTd
		CloseTr
		OpenTr "teven"
			OpenTd "teven", "align=left "
				Response.Write "Solicitud de"
			CloseTd
			OpenTd "teven", "align=left width=500px "
				Response.Write arrRadicationDetails(0,0)
			CloseTd
		CloseTr
		OpenTr "todd"
			OpenTd "todd", "align=left "
				Response.Write "Solicitud No."
			CloseTd
			OpenTd "todd", "align=left width=500px "
				Response.Write arrRadicationDetails(1,0)
			CloseTd
		CloseTr

		OpenTr "teven"
			OpenTd "teven", "align=left "
				Response.Write "Contrato"
			CloseTd
			OpenTd "teven", "align=left width=500px"
				'Response.Write arrRadicationDetails(2,0) 
				'==========================================================================
				'<Alejandro Jaramillo  2009/06/04 - Modificado por proceso de Referencia Unica - Recaudos>
				' Obtiene la referencia única a partir del Contrato y el producto
				'==========================================================================
				If (Trim(arrRadicationDetails(4,0)) = "FCO") Then 
					Reference = GetReferenciaUnica(Trim(arrRadicationDetails(4,0)), Trim(arrRadicationDetails(3,0)) & Trim(arrRadicationDetails(2,0)))
				Else
					Reference = GetReferenciaUnica(Trim(arrRadicationDetails(4,0)), Trim(arrRadicationDetails(2,0)))
				End if
				'==========================================================================
				if IsNull(Reference) or len(Reference)>12 Then
					Response.Write arrRadicationDetails(2,0)
				else
					Response.Write Reference
				end if
			CloseTd
		CloseTr
		OpenTr "todd"
			OpenTd "todd", "align=left "
				Response.Write "Plan"
			CloseTd
			OpenTd "todd", "align=left width=500px"
				Response.Write arrRadicationDetails(3,0)
			CloseTd
		CloseTr

		OpenTr "teven"
			OpenTd "teven", "align=left "
				Response.Write "Producto"
			CloseTd
			OpenTd "teven", "align=left width=500px"
				Response.Write arrRadicationDetails(4,0)
			CloseTd
		CloseTr
		OpenTr "todd"
			OpenTd "todd", "align=left "
				Response.Write "Fecha Recepción"
			CloseTd
			OpenTd "todd", "align=left width=500px"
				Response.Write year(arrRadicationDetails(5,0)) &"/"& _
						Month(arrRadicationDetails(5,0)) &"/"& _
						day(arrRadicationDetails(5,0))
			CloseTd
		CloseTr

		OpenTr "teven"
			OpenTd "teven", "align=left "
				Response.Write "Usuario receptor"
			CloseTd
			OpenTd "teven", "align=left width=500px"
				Response.Write arrRadicationDetails(6,0)
			CloseTd
		CloseTr
		OpenTr "todd"
			OpenTd "todd", "align=left "
				Response.Write "Forma de recepción"
			CloseTd
			OpenTd "todd", "align=left width=500px"
				Response.Write arrRadicationDetails(7,0)
			CloseTd
		CloseTr

		OpenTr "teven"
			OpenTd "teven", "align=left "
				Response.Write "Estado"
			CloseTd
			OpenTd "teven", "align=left width=500px"
				Response.Write arrRadicationDetails(8,0)
			CloseTd
		CloseTr	
		OpenTr "todd"
			OpenTd "todd", "align=left "
				Response.Write "Comentario"
			CloseTd
			OpenTd "todd", "align=left width=500px"
				Response.Write arrRadicationDetails(9,0)
			CloseTd
		CloseTr	

		OpenTr "teven"
			OpenTd "teven", "align=left "
				Response.Write "Valor"
			CloseTd
			OpenTd "teven", "align=left width=500px"
				If Not(IsNull(arrRadicationDetails(11,0)))  then
					Response.Write FormatCurrency(arrRadicationDetails(11,0))
				End if
			CloseTd
		CloseTr	
		OpenTr "todd"
			OpenTd "todd", "align=left "
				Response.Write "Tipo de retiro"
			CloseTd
			OpenTd "todd", "align=left width=500px"
				Response.Write arrRadicationDetails(12,0)
			CloseTd
		CloseTr	
	CloseTable

OpenTable "50%", "'' align=center border=0"
	OpenTr "teven"
		OpenTd "todd", "align=left "
		


if Clng(arrRadication(0,0)) <> clng(arrRadication(6,0)) then
	OpenForm "post10Up","Post", "radicationlist.asp", ""		
				PlaceInput "LastRadID", "hidden", arrRadication(0,UBound(arrRadication,2)), ""
				PlaceInput "FirstRadID", "hidden", arrRadication(0,0), ""
				PlaceInput "UpDown", "hidden", "Up", ""
				PlaceInput "RadDetails", "hidden", "" , ""								
				PlaceInput "Contract", "hidden", Contract, ""
				PlaceInput "Product", "hidden", Product, ""
				PlaceInput "Plan", "hidden", Plan, ""
				PlaceInput "ClientId", "hidden", nroDocum, ""
				PlaceInput "DocType", "hidden", DocType, ""
				PlaceInput "Name", "hidden", Name, ""
				PlaceInput "Back", "Submit", "Anteriores 10", "class=sbttn"		
	CloseForm
end if
		CloseTd
		OpenTd "todd", "align=left "
if clng( arrRadication(0,Ubound(arrRadication,2 ))) <> clng(arrRadication(7,0)) then
	OpenForm "post10Up","Post", "radicationlist.asp", ""		
				PlaceInput "LastRadID", "hidden", arrRadication(0,UBound(arrRadication,2)), ""
				PlaceInput "FirstRadID", "hidden",  arrRadication(0,0), ""
				PlaceInput "UpDown", "hidden", "Down", ""
				PlaceInput "RadDetails", "hidden", "" , ""								
				PlaceInput "Contract", "hidden", Contract, ""
				PlaceInput "Product", "hidden", Product, ""
				PlaceInput "Plan", "hidden", Plan, ""
				PlaceInput "ClientId", "hidden", nroDocum, ""
				PlaceInput "DocType", "hidden", DocType, ""
				PlaceInput "Name", "hidden", Name, ""
				PlaceInput "Back", "Submit", "Siguientes 10", "class=sbttn"		
	CloseForm
end if	
		CloseTd
	CloseTr
CloseTable
	OpenTable "70%", "'' align=center border=0"
		OpenTr ""
			OpenTd "thead", "align=center "
				Response.Write "Solicitud"
			CloseTd
			OpenTd "thead", "align=center "
				Response.Write "Contrato"
			CloseTd
			OpenTd "thead", "align=center "
				Response.Write "Plan"
			CloseTd
			OpenTd "thead", "align=center "
				Response.Write "Producto"
			CloseTd
			OpenTd "thead", "align=center "
				Response.Write "Fecha"
			CloseTd
			OpenTd "thead", "align=center "
				Response.Write "Proceso"
			CloseTd
			OpenTd "thead", "align=center "
				Response.Write "Detalles"
			CloseTd
		CloseTr

'== START customer Name ==

dim errorCatcher
Err.Clear
for J = 0  to UBound(arrRadication,2)
		
		'on error resume next
			errorCatcher = arrRadication(0,J)
			If Err.number <> 0 Then
				exit for
			end if
		if CDbl(arrRadication(0,J)) = cdbl( arrRadicationDetails(1,0)) then
			classid = "tdetailed"
		else
			if (2 * Round(J / 2)) = J then
				classid = "teven"
			else
				classid = "todd"
			end if
		end if
		OpenTr ""
			OpenTd classid, "align=center "
				Response.Write arrRadication(0,J)				
			CloseTd
			OpenTd classid, "align=center "
				'Response.Write arrRadication(1,J) & "<br>"
				'<Alejandro Jaramillo  2009/06/04 - Modificado por proceso de Referencia Unica - Recaudos>
				If (trim(Trim(arrRadication(3,J))) = "FCO") Then 
					Reference = GetReferenciaUnica(Trim(arrRadication(3,J)), Trim(arrRadication(2,J)) & Trim(arrRadication(1,J)))
				Else
					Reference = GetReferenciaUnica(Trim(arrRadication(3,J)), trim(arrRadication(1,J)))
				End if
				if IsNull(Reference) or len(Reference)>12 Then
					Response.Write arrRadication(1,J)
				else
					Response.Write Reference
				end if
			CloseTd
			OpenTd classid, "align=center "
				Response.Write arrRadication(2,J)
			CloseTd
			OpenTd classid, "align=center "
				Response.Write arrRadication(3,J)
			CloseTd
			OpenTd classid, "align=center "
				Response.Write arrRadication(4,J)
			CloseTd
			OpenTd classid, "align=center "
				Response.Write arrRadication(5,J)
			CloseTd
			OpenTd classid, "align=center "

	OpenForm "poste"&cstr(J),"Post", "radicationlist.asp", ""
				PlaceInput "LastRadID", "hidden", Request.Form("LastRadID"), ""
				PlaceInput "FirstRadID", "hidden", Request.Form("FirstRadID"), ""
				PlaceInput "UpDown", "hidden", Request.Form("UpDown"), ""
				PlaceInput "RadDetails", "hidden", arrRadication(0,J) , ""								
				PlaceInput "Contract", "hidden", Contract, ""
				PlaceInput "Product", "hidden", Product, ""
				PlaceInput "Plan", "hidden", Plan, ""
				PlaceInput "ClientId", "hidden", nroDocum, ""
				PlaceInput "DocType", "hidden", DocType, ""
				PlaceInput "Name", "hidden", Name, ""
				PlaceInput "Back", "Submit", "Detalles", "class=sbttn"		
	CloseForm
				
			CloseTd
		CloseTr
next
CloseTable
	OpenTable "70%", "'' align=center border=0"
		OpenTr ""
			OpenTd "thead", "align=center "
				Response.Write "Solicitudes en pantalla " & _
				UBound(arrRadication,2) + 1
			CloseTd
	CloseTable
CloseBody
CloseHTML
%>
