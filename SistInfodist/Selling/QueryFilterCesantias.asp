<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:					QueryFilterCesantias.asp  19307
'Path:						/sistinfodist/Selling/QueryFilterCesantias.asp
'Created By:				Margarita Cardozo 2005/10/15
'Last Modified:				Diana Mariced Pérez	2008/05/12	Added PlaceTitle and changed page_id in write_sp_log		
'Parameters:						
'Returns:						
'Additional Information:	filter
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
Dim user
Dim Sql,strSQL 'SQL Sentences holder
Dim adoConn,oConn'Database Connection
Dim bc
Dim rstcampo
Dim arrProductos, arrSociedades, arrAgente, arrWsaler
Dim rs, cn,objRst,sel,val,nivel_detalle
Dim I,J,K,L      
Dim Idsociedad, Idagte, Wsaler
Dim Page
dim FechaCorte
Dim mesIni, anoIni, mesfin, anoFin
dim nc
dim parametros
dim CodFPOB1
dim conexion
'======================================================================================
'Check Browser
'======================================================================================
Set bc = Server.CreateObject("MSWC.BrowserType")
Authorize 7,24
Set adoConn = GetConnPipelineDB
'write_sp_log adoConn, 19307, "", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin")
'write_sp_log adoConn, 19307, "", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin") & " acclev " & Cstr(Session("sp_AccessLevel"))
write_sp_log adoConn, 13358, "", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin")
write_sp_log adoConn, 13358, "", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin") & " acclev " & Cstr(Session("sp_AccessLevel"))
CloseConnPipelineDB
Set adoConn = Nothing


''INICIALIZAR VARIABLES

''''inicializar para combos de acuerdo al nivel de acceso
dim AccessLevel, idAgteLoggedIn, idSociedadLoggedIn, idAgteContract, idSocContract 
dim isAuthorized 
AccessLevel= Cstr(Session("sp_AccessLevel"))
idAgteLoggedIn= CStr(Session("sp_IdAgte"))
idSociedadLoggedIn= CStr(Session("sp_Idsoc"))

wsaler=""
idagte=idAgteLoggedIn
idsociedad=idSociedadLoggedIn

nivel_detalle=""
mesini=month(now())
anoini=year(now())
CodFPOB1=Request.Form ("CODFPOBquery1")
isAuthorized  = false

conexion=""

Select Case AccessLevel
		Case 0 'Skandia
				wsaler=null
				idagte=0
				idsociedad=0
				isAuthorized  = true		
		Case 1 'WHOLE SALER	
				wsaler = get_wsalerwsaler(Session("sp_Idworker"))
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
	
if not isAuthorized then
		Response.Write("Acceso no autorizado")	
		Response.Redirect Application("UnauthorizedURL")		  
end if
'-------------------------------------------------------------
get_agente  idagte, idsociedad,WSALER
get_sociedad idsociedad,wsaler
write_dataLog Response.Status,"QueryFilterCesantias.asp", "QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),conexion,"N/A","null","Consulta","N/A"

function get_wsalerwsaler( idworker )
		dim rstcampo1
		dim adoconn1
		dim arrwsaler
		Set adoConn1 = GetConnpipelineDB()		
		Set rstCampo1 = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.wsalerGetWsaler " & cdbl(idworker)
		conexion=Sql&" - "
		
		rstCampo1.Open Sql, adoConn1
		If rstCampo1.BOF And rstCampo1.EOF Then
			get_wsalerwsaler = null
		Else
			arrwsaler = rstCampo1.GetRows()
			get_wsalerwsaler = arrwsaler(0,0)
		End If
		rstCampo1.Close
		adoConn1.close
		Set rstCampo1 = nothing
		set adoconn1=nothing
end function


function get_sociedad( idsociedad,WSALER )
		dim sql1
		Set adoConn = GetConnpipelineDB()		
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.spcm_getsociedadfp " & cdbl(idsociedad) & ",'" & WSALER & "'"
		Sql1= replace(Sql,"'","")
		conexion= conexion&"exec sigscg.dbo.spcm_getsociedadfp " & cdbl(idsociedad) & " -" & WSALER&"-"
		'write_sp_log adoConn, 19307, "", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin") & " sql " & sql1
		write_sp_log adoConn, 13358, "sigscg.dbo.spcm_getsociedadfp", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin") & " sql " & sql1
		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			arrSociedades = 0
		Else
			arrSociedades = rstCampo.GetRows()
		End If
		rstCampo.Close
		adoConn.close
		Set rstCampo = nothing
		set adoconn=nothing
end function


function get_agente(par_idagte, par_idsociedad,  par_WSALER )
		dim sql1
		Set adoConn = GetConnpipelineDB()
		Set rstCampo = Server.CreateObject("ADODB.Recordset")

		' Armando Arias - 01/Ago/2008 - Se cambia el nombre del agente por "ExternalReference" de la misma tabla
		' Sql = "exec sigscg.dbo.spcm_getagente " & cdbl(par_idagte) & "," & cdbl(par_idsociedad) & ",'" & par_WSALER & "'"
		Sql = "exec sigscg.dbo.spcm_getagente_nominado " & cdbl(par_idagte) & "," & cdbl(par_idsociedad) & ",'" & par_WSALER & "'"

		Sql1= replace(Sql,"'","")
		conexion= conexion&"exec sigscg.dbo.spcm_getagente_nominado " & cdbl(par_idagte) & "-" & cdbl(par_idsociedad) & "-" & par_WSALER &"-"
		'write_sp_log adoConn, 19307, "", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin") & " sql " & sql1
		write_sp_log adoConn, 13358, "", 0, "", "", 0, 0, "", "/sistinfodist/Selling/QueryFilterCesantias.asp Loaded by " & Session("sp_miLogin") & " sql " & sql1

		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			arrAgente = 0
			
		Else

			arrAgente = rstCampo.GetRows()
		End If
		
		rstCampo.Close
		adoConn.close
		Set rstCampo = nothing
		set adoconn=nothing
end function

OpenHTML

OpenHead
PlaceTitle "QueryFilterCesantias"

PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
PlaceMeta "Pragma", "", "no_cache"


			

%><script language="javascript" src="/super_pipeline/operations/_pipeline_scripts/SkCoSecurity.js"></script>
<%

%>
<SCRIPT LANGUAGE=javascript ></SCRIPT>

<SCRIPT LANGUAGE=javascript>




//'array sociedades lado cliente
//fila volteo el arreglo
var Sociedades = new Array(<%=UBound(arrSociedades,2)%>)
for (i = 0; i <= <%=UBound(arrSociedades,2)%>; i ++) {
	Sociedades[i] = new Array(<%=UBound(arrSociedades)%>)
}
<%
For i = 0 To UBound(arrSociedades,2)
	For j = 0 To UBound(arrSociedades)
		'If( Not(IsNull(arrSociedades(I, J)))) Then
	'	If( j=3 ) Then
	'		arrSociedades(I, J) = Replace(arrSociedades(I, J),vbCrLf,"")
	'	End If
		Response.Write "Sociedades[" & I & "][" & J & "] = '" & arrSociedades(j,i) & "'" & vbCrLf
	Next
Next

%>

//'array agentes lado cliente

var Agentes = new Array(<%=UBound(arrAgente, 2)%>)
for (i = 0; i <= <%=UBound(arrAgente, 2)%>; i ++) {
	Agentes[i] = new Array(<%=UBound(arrAgente)%>)
}
<%
For I = 0 To UBound(arrAgente, 2)
	For J = 0 To UBound(arrAgente)
		If J = 3 And Not(IsNull(arrAgente(I, J))) Then
			arrAgente(I, J) = Replace(arrAgente(I, J),vbCrLf,"")
		End If
		Response.Write "Agentes[" & I & "][" & J & "] = '" & arrAgente(J, I) & "'" & vbCrLf
	Next
Next
%>


function addoption(id, name,tbox)
{
	var no = new Option();
	no.value = id;
	no.text = name;
    tbox.options[tbox.options.length] = no;
}


function fillFP(obj) 
{
		var Fp = obj.selectedIndex;
	    document.frProducto.idAgte0.length = 1
		for (i = 0; i < Agentes.length; i ++) 
		{ 
		
			if (Agentes[i][2] == Sociedades[Fp][1]) 
				{   
				     addoption(Agentes[i][0], Agentes[i][1],document.frProducto.idAgte0)
				}
		}
	return true;
}

function back()
{
	frProducto.action="default.asp";
	document.frProducto.submit();
	
}

function ChangeidAgte1()
{
var selected ;

}


function go()

{	
	
	var selected ;
	var cancel	=	true;
	//if ( (document.frProducto.idsociedad0.selectedIndex==null ) || (document.frProducto.idsociedad0.selectedIndex==-1) ||document.frProducto.idsociedad0.selectedIndex==0 )

	if ( document.frProducto.CodFPOB.value	==null  || document.frProducto.CodFPOB.value=="" ) 
	{
			alert ("Por digite el codigo del agente en Cesantias");
			
	}
	
	
	if ( (document.frProducto.idsociedad0.selectedIndex==null ) || (document.frProducto.idsociedad0.selectedIndex==-1)  )
	{
			document.frProducto.pagina1.value="./queryfiltercesantias.asp";
			alert ("Por favor elija una sociedad valida");
			
	}

	else
	{
		if ( (document.frProducto.idAgte0.selectedIndex==null ) || (document.frProducto.idAgte0.selectedIndex==-1) ||document.frProducto.idAgte0.selectedIndex==0 )
		{
				document.frProducto.pagina1.value="./queryfiltercesantias.asp";
				alert ("Por favor elija un agente válido");
		}

		else
		{
	
				var ag=document.frProducto.idAgte0.options[document.frProducto.idAgte0.selectedIndex].value;	
				var so=document.frProducto.idsociedad0.options[document.frProducto.idsociedad0.selectedIndex].value;
				
				//especifico para cesantias
				document.frProducto.nivel_detalle1.value	='AG';
				document.frProducto.IdAgtequery1.value		=	ag;
				document.frProducto.CODFPOBquery1.value		=	document.frProducto.CodFPOB.value;
			
				//para la siguiente pagina						
				document.frProducto.idagte1.value = ag;
				document.frProducto.nombreagte1.value 		= document.frProducto.idAgte0.options[document.frProducto.idAgte0.selectedIndex].text;
	
				document.frProducto.idsociedad1.value		= document.frProducto.idsociedad0.options[document.frProducto.idsociedad0.selectedIndex].value;
				document.frProducto.nombresoc1.value		= document.frProducto.idsociedad0.options[document.frProducto.idsociedad0.selectedIndex].text;	
					
				document.frProducto.pagina0.value			="./EmployerResultsbyFP.asp";
				document.frProducto.pagina1.value			="./EmployerResultsbyFP.asp";
				cancel = false;
				frProducto.action=document.frProducto.pagina1.value;		
		}
	}
	
	if (cancel == false)
		{
			document.frProducto.submit();
			}
	return true;
	
}



</SCRIPT>
<%
	PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	PlaceLink "REL", "stylesheet", "../../css/style1.css", "text/css"

closehead

OpenBody "", ""
OpenTable "90%", " align=center cellpadding=0 cellspacing=0"
	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
	
		CloseTd
	CloseTr
	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
			Response.Write "<h3> Consulta Empresas Cesantias </h3>"
		CloseTd
	CloseTr
	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
	
		CloseTd
	CloseTr
	
	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
			Response.Write "<hr>"
		CloseTd
	CloseTr

	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
			Response.Write "&nbsp;"
		CloseTd
	CloseTr
	
closetable	
	
If autorizarMn(7,24) Then

OpenForm "frProducto", "post", "../default.asp", ""
'OpenForm "frProducto", "post", "", ""
OpenTable "75%", "'' align=center"			
OpenTr "valign=top"
	OpenTd "thead", "width=20% align=left"
		Response.Write "Sociedad" 
	CloseTd
	
	    
	OpenTd "", ""
	    OpenCombo "idsociedad0", " class=bttntext onclick='javascript:fillFP(this)' id='Sociedad'"
	    	For J = 0 To UBound(arrsociedades, 2)
				If CStr(Request.Form("idSociedad1")) = CStr(arrsociedades(0,J)) Then
					Sel = "selected"
				Else
					Sel = ""
				End If
				PlaceItem Sel, arrsociedades(0,J), arrsociedades(1,J)
			Next 'J
		CloseCombo
	CloseTd

CloseTr



OpenTable "75%", "'' align=center"	
	OpenTr "class=teven"
		OpenTd "thead", "width=20% align=left"
		Response.Write "Agente" 
		
	CloseTd
	OpenTd "tbody", "width=90% align=left"
	    idsociedad = Request.Form("idSociedad1")
		get_agente idagte, idsociedad, wsaler
		
		%><SELECT NAME="idAgte0"  class="bttntext" onChange="ChangeidAgte1();"><%		
		%>	</SELECT>	<%
	CloseTd

CloseTr

	
CLOSETABLE		

OpenTable "75%", "'' align=center"	
	OpenTr "class=teven"
		OpenTd "thead", "width=20% align=left"
			Response.Write "PROM Cesantias" 
		CloseTd
		OpenTd "teven", "align=left "
			PlaceInput "CodFPOB", "Input", CodFPOB1 ,""
		CloseTd
	
	CloseTr
	
	OpenTr "class=teven"
		
		

	CloseTr

	
CLOSETABLE		


OpenTable "75%", "'' align=center height=14"
	OpenTr "class=thead"

		OpenTd "tbody2", "width=10% align=left"
			Response.Write "&nbsp;"
		CloseTd
		OpenTd "tbody2", "width=25% align=left"
				PlaceInput "Enviar ", "submit", "Enviar ", "class=sbttn  onclick=go()"
				PlaceInput "Regresar", "submit", "Regresar","class=sbttn   onclick=back()"
		CloseTd

		OpenTd "tbody2", "width=30% align=left"
			Response.Write "&nbsp;"
		CloseTd
		
	CloseTr
	OpenTr "class=todd"
		PlaceInput "nivel_detalle1","hidden","",""
		PlaceInput "parametros1","hidden","",""
		PlaceInput "nombre_parametro1","hidden","",""
		PlaceInput "Accesslevel","hidden","",""
		PlaceInput "IdAgtequery1","hidden","",""
		PlaceInput "CODFPOBquery1","hidden","",request.Form("CODFPOBquery1")
		
		PlaceInput "nombresoc1","hidden",Request.Form ("nombresoc1"),""
		PlaceInput "nombreagte1","hidden",Request.Form ("nombreagte1"),""
		
		
		PlaceInput "idsociedad1","hidden",Request.Form ("idsociedad1"),""
		PlaceInput "idagte1","hidden",Request.Form ("idagte1"),""

		PlaceInput "pagina1","hidden","../default.asp",""'por defecto va a default
		PlaceInput "pagina0","hidden","./queryfiltercesantias.asp",""
		PlaceInput "pagina_2","hidden","../default.asp",""
	CloseTr


CloseForm
end if
CloseBody
CloseHTML

Response.End 
If Err.number <> 0 Then'
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>

