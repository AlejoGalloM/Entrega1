<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		searchEmployee.asp 19301
'Path:				search/
'Created By:		Margarita Cardozo
'					Jimmy Ospino
'Last Modified:		2003/07/30
'Parameters:		    Identification Number 
'						Name 
'	Esta pagina no tiene bloqueo de boton derecho del mouse por requerimiento
'	
'Returns:			
'===================================================================================
Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>

<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../../operations/_pipeline_scripts/url_check.asp"-->
<%
Authorize 1,25
Dim adoConn
Dim objRst
Dim strSQL
Dim Combo
Dim ComboDocType
Dim Contract, ClientId, Name, LastName
Dim Init_msg
Dim Idsociedad, Idagte, Wsaler
Dim Page 
Dim conexion

dim sql1,sql
sql1="searchemployee.asp"
sql1=replace(sql1,"'","")
Set adoConn = GetConnpipelineDB()
write_sp_log adoConn, 19301, MID (sql1,1,50), 0, "", "", 0, 0, "", mid( "sistinfodist/selling/searchEmployee.asp Loaded by " & Session("sp_miLogin") & "- par: " & sql1 ,1,250)
CloseConnPipelineDB

conexion=""

''''inicializar para combos de acuerdo al nivel de acceso
dim AccessLevel, idAgteLoggedIn, idSociedadLoggedIn, idAgteContract, idSocContract 
dim isAuthorized 

AccessLevel= Cstr(Session("sp_AccessLevel"))
idAgteLoggedIn= CStr(Session("sp_IdAgte"))
idSociedadLoggedIn= CStr(Session("sp_Idsoc"))
wsaler=""
idagte=idAgteLoggedIn
idsociedad=idSociedadLoggedIn

isAuthorized  = false

Select Case AccessLevel
		Case 0 'Skandia
				wsaler=null
				idagte=0
				idsociedad=0
				isAuthorized  = true		
		Case 1 'WHOLE SALER	
				wsaler = get_wsalerwsaler(Session("sp_Idworker"))
				conexion= "exec sigscg.dbo.WsalerGetWsaler " & idworker
				if not isnull(wsaler) then
					isAuthorized = true			
				end if	
		Case 2 'Partner FP, Fran. Worker
				idagte=0
				idsociedad=idSociedadLoggedIn
				isAuthorized = true
		Case 3 'FP
				idagte=idAgteLoggedIn
				idsociedad=idSociedadLoggedIn
				isAuthorized = true
		case else
				isAuthorized = false
		End Select

	write_dataLog Response.Status,"searchEmployee.asp", "searchEmployee.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),conexion,"N/A","null","Consulta","N/A"
	
if not isAuthorized then
		Response.Write("Acceso no autorizado")	
		Response.Redirect Application("UnauthorizedURL")		  
end if
'-------------------------------------------------------------
function get_WsalerWsaler( idworker )
		dim rstcampo1
		dim adoconn1
		dim arrWsaler
		Set adoConn1 = GetConnpipelineDB()		
		Set rstCampo1 = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.WsalerGetWsaler " & idworker
		
		rstCampo1.Open Sql, adoConn1
		If rstCampo1.BOF And rstCampo1.EOF Then
			get_WsalerWsaler = null
		Else
			arrWsaler = rstCampo1.GetRows()
			get_WsalerWsaler = arrWsaler(0,0)
		End If
		rstCampo1.Close
		adoConn1.close
		Set rstCampo1 = nothing
		set adoconn1=nothing
end function


OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src='/super_pipeline/operations/_pipeline_scripts/validation.js'></SCRIPT>
<%
PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
PlaceLink "REL", "stylesheet", "../../css/style1.css", "text/css"
CloseHead
%>
<SCRIPT>


function  validar(obj)
	{

	if ( 
	(document.searchNames.NombreEmp1.value == null || document.searchNames.NombreEmp1.value == "" )
	&&
	(document.searchClientId.NroDocumEmp1.value == null || document.searchClientId.NroDocumEmp1.value == "" )
	 )
		{
				alert("Por favor complete el campo Nombre de la Empresa o Número de Documento");
				document.searchNames.action="./searchemployee.asp";
				document.searchClientId.action="./searchemployee.asp";
				return;
		}
	else 	
		{
		return formValidation(obj);
		}

}

function go1 ()
{
  
	frcover.action="default.asp"
	document.frCover.submit();
}

</SCRIPT>
<%


OpenBody "''", ""
	OpenTable "90%", "'' align=center cellpadding=0 cellspacing=0 border=0"
	
		OpenTr "valign=middle"
			OpenTd "''", "align=center valign=middle "
				Response.Write "<h3>Consulta Empresa </h3>"
					
			CloseTd
		CloseTr
		
		OpenTr "valign=middle"
			OpenTd "''", "align=center valign=middle "
				Response.Write "<hr>"
			CloseTd
		CloseTr
	closetable	
		
	OpenTable "75%", "'' align=center"

		OpenTr ""
			OpenTd "teven", "colspan=4"
				Response.Write "&nbsp;"
			CloseTd					  
		CloseTr
	
		OpenForm "searchNames", "post", "searchEmployee_results.asp", "onSubmit='return validar(this)'"
		OpenTr ""
			OpenTd "thead", "align=left"
				PlaceInput "btnClientName", "submit", "   Buscar x Nombre   ", "id='               Enviar' class=sbttn"
			CloseTd
			OpenTd "tbody", "colspan=2"
				PlaceInput "NombreEmp1", "text", "", "id='             L  Nombres' class=bttntext"
				PlaceInput "sp", "hidden", "spsp_SearchNames", "id='               SP'"
			CloseTd
		CloseTr
		CloseForm
		
	OpenTr ""
		OpenTd "teven", "colspan=4"
			Response.Write "&nbsp;"
		CloseTd					  
	CloseTr
	
		
	OpenForm "searchClientId", "post", "searchEmployee_results.asp", "onSubmit='return validar(this)'"
	OpenTr ""
		OpenTd "thead", "align=left nowrap"
			PlaceInput "btnClientId", "submit", "Buscar x NIT ", "id='               Enviar' class=sbttn"
		CloseTd
		OpenTd "tbody", ""
			PlaceInput "NroDocumEmp1", "text", "", "id='RN       I     Número Identificación' class=bttntext"
			
		CloseTd

		OpenTd "", ""
			Response.Write "&nbsp;"
		CloseTd					  				
	CloseTr
	CloseForm
	OpenTr ""
		OpenTd "teven", "colspan=4"
			Response.Write "&nbsp;"
		CloseTd					  
	CloseTr

	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
			Response.Write "&nbsp;"
		CloseTd
	CloseTr
CloseTable

CloseTable
OpenTable "95%", "'' align=center border=0 height=1"
	OpenTr "class=tbody2"
		OpenTd "''","align=center"
		OpenForm "cons", "post", "../selling/default.asp", "onSubmit='javascript:return formValidation(this)'"
				PlaceInput "Regresar", "submit", "Volver","class=sbttn  onclick='';"
		closeform		 
		CloseTd
	closetr
closetable	

CloseBody
CloseHTML

If Err.number <> 0 Then

	Set bc = Server.CreateObject("MSWC.BrowserType")
	Set adoConn = GetConnpipelineDB()
	write_sp_log adoConn, 19305, "Error", 0, "", "", 0, 0, "", mid( "ClientContractDetails.asp Loaded by " & Session("sp_miLogin") & " err:" & err.Description ,1,250)
	CloseConnPipelineDB

	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>