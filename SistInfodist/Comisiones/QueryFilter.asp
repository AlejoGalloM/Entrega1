<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:			 queryfilter.asp  19101
'Path:				 /sistinfodist/comisiones/queryfilter.asp
'Created By:			 Margarita Cardozo 2003/05/26
'Last Modified:			 mmc 2003/08/22   mmc 2003/08/27
'                                Armando Arias - 2008/May/12 - Cumplimiento circular SF052 - PlaceTitle - write_sp_log 13348
'Parameters:
'Returns:
'Additional Information:	filter
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
Dim fecFinalConsultaComision
dim nc
dim parametros

'======================================================================================
'Check Browser
'======================================================================================
Authorize 7,24
Set adoConn = GetConnPipelineDB

'write_sp_log adoConn, 19101, "", 0, "", "", 0, 0, "", "/sistinfodist/comisiones/queryfilter.asp Loaded by " & Session("sp_miLogin")
write_sp_log adoConn, 13348, "", 0, "", "", 0, 0, "", "/sistinfodist/comisiones/queryfilter.asp Loaded by -" & Session("sp_miLogin")

CloseConnPipelineDB
Set adoConn = Nothing

''INICIALIZAR VARIABLES

'======================================================================================
'security
'======================================================================================

dim todosag, todosso
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

isAuthorized  = false


'Response.Write AccessLevel
Select Case AccessLevel
		Case 0 'Skandia
				wsaler=null
				idagte=0
				idsociedad=0
				todosso=0
				todosag=1
				isAuthorized  = true
		Case 1 'WHOLE SALER
				wsaler = get_wsalerwsaler(Session("sp_Idworker"))
				if not isnull(wsaler) and wsaler<>"" and wsaler <>"0" then
					isAuthorized = true
				end if
				todosso=0
				todosag=1
		Case 2 'Partner FP, Fran. Worker
				idagte=0
				idsociedad=idSociedadLoggedIn
				if idsociedad>0 then
					isAuthorized = true
				end if
				todosso=0
				todosag=1
		Case 3 'FP
				idagte=idagteLoggedIn
				idsociedad=idSociedadLoggedIn
				if idsociedad>0 and idagte>0 then
					isAuthorized = true
				end if
				todosso=0
				todosag=0

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

function get_wsalerwsaler( idworker )
		dim rstcampo1
		dim adoconn1
		dim arrwsaler
		Set adoConn1 = GetConnpipelineDB()
		Set rstCampo1 = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.wsalerGetWsaler " & idworker

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
		Set adoConn = GetConnpipelineDB()
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec Comision.dbo.spcm_getsociedadfp " & idsociedad & ",'" & WSALER & "'"

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


function get_agente(idagte, idsociedad,WSALER )
		Set adoConn = GetConnpipelineDB()
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		' Armando Arias - 01/Ago/2008 - Se cambia el nombre del agente por "ExternalReference" de la misma tabla
		Sql = "exec sigscg.dbo.spcm_getagente " & idagte & "," & idsociedad & ",'" & WSALER & "'"
		'sql = "exec COMISION.dbo.spcm_getagente_nominado " & idagte & "," & idsociedad & ",'" & WSALER & "'"

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
      PlaceTitle "QueryFilter"
      PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
      PlaceMeta "Pragma", "", "no_cache"

%>

<SCRIPT language="javascript" src="/super_pipeline/operations/_pipeline_scripts/SkCoSecurity.js"></script>
<SCRIPT LANGUAGE=javascript>

//'array sociedades lado cliente
//var Sociedades = new Array(<%=UBound(arrSociedades, 2)+2 %>)  //dim tamaño de filas

var largo=<%=UBound(arrSociedades, 2) + 2 %>

var Sociedades = new Array(largo); //dim tamaño de filas
//aqui voy a adicionar todas las sociedades
for (i = 0; i < largo; i ++) {  //por eso arranca en 1
	Sociedades[i] = new Array(<%=UBound(arrSociedades,1)%>);
}
<%
'if todosso=1 then
'	Response.Write"Sociedades[0][0] = ' '" & vbCrLf
'	Response.Write"Sociedades[0][1] = 'Todas las sociedades'" & vbCrLf
'	Response.Write"Sociedades[0][2] = ''" & vbCrLf
'	Response.Write"Sociedades[0][3] = ''" & vbCrLf
'end if

For J = 0 To UBound(arrSociedades, 2) 'filas
	For I = 0 To UBound(arrSociedades) 'col
		If I = 3 And Not(IsNull(arrSociedades(I, J))) Then
			arrSociedades(I, J) = Replace(arrSociedades(I, J),vbCrLf,"")
		End If
		if todosso=1 then
			Response.Write"Sociedades[" & j+1 & "][" & i & "] = '" & arrSociedades(i, j) & "'" & vbCrLf
		else
			Response.Write"Sociedades[" & j & "][" & i & "] = '" & arrSociedades(i, j) & "'" & vbCrLf
		end if
	Next
Next
%>



//'array agentes lado cliente
largo=<%=UBound(arrAgente, 2) + 2 %>
var agentes = new Array(largo);
for (i = 0; i < largo; i ++) {
		agentes[i] = new Array(<%=UBound(arrAgente,1)%>);
	}
<%
if todosag=1 then
		Response.Write"agentes[0][0] = ' '" & vbCrLf
		Response.Write"agentes[0][1] = '-- Cover Sociedad --'" & vbCrLf
		Response.Write"agentes[0][2] = ''" & vbCrLf
end if
For J = 0 To UBound(arrAgente, 2) 'fila
	For I = 0 To UBound(arrAgente,1)  'col
		If I = 3 And Not(IsNull(arrAgente(i, j))) Then
			arrAgente(I, J) = Replace(arrAgente(I, J),vbCrLf,"")
		End If
		if todosag=1 then
			Response.Write"agentes[" & j+1 & "][" & i & "] = '" & arrAgente(i,j) & "'" & vbCrLf
		else
			Response.Write"agentes[" & j & "][" & i & "] = '" & arrAgente(i,j) & "'" & vbCrLf
		end if
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
	    document.frProducto.idAgte1.length = 1;
	  //  alert (String(agentes.length));
		for (i = 0; i < agentes.length; i ++)
		{
			if (agentes[i][2] == Sociedades[Fp][1])
				{
				     addoption(agentes[i][0], agentes[i][1],document.frProducto.idAgte1);
				}
		}
	return true;


}

function back()
{
	document.frProducto.action="../default.asp";
	document.frProducto.submit();

}
//go
function go(pAno, pMes)
{
	var selected ;
	var as; // Año final seleccionado en el combo
	var ms; // Mes final seleccionado en el combo
	var meses 
	
	switch(pMes)
	{
		case  1 : meses = "Enero";
				 break;	
		case  2 : meses = "Febrero";
				 break;	
		case  3 : meses = "Marzo";
				 break;	
		case  4 : meses = "Abril";
				 break;	
		case  5 : meses = "Mayo";
				 break;	
		case  6 : meses = "Junio";
				 break;	
		case  7 : meses = "Julio";
				 break;	
		case  8 : meses = "Agosto";
				 break;	
		case  9 : meses = "Septiembre";
				 break;	
		case 10 : meses = "Octubre";
				 break;	
		case 11 : meses = "Noviembre";
				 break;	
		case 12 : meses = "Diciembre";
				 break;	
	} 

	as = document.frProducto.toyear1.value
	ms = document.frProducto.tomonth1.value
	
	if ( (document.frProducto.idAgte1.selectedIndex==null ) || (document.frProducto.idAgte1.selectedIndex==-1) )
	{
			document.frProducto.pagina1.value="../default.asp";
			alert ("Por favor elija un agente válido");
	}
	else if (as > pAno)
	{
			//document.frProducto.pagina1.value="../default.asp";
			alert ("No puede seleccionar un año final superior a " + af);
			return false;
	}
	else if ((as == pAno) && (ms > pMes))
	{
			//document.frProducto.pagina1.value="../default.asp";
			alert ("Puede consultar hasta " + meses + " del " + pAno);
			return false;
	}
	else
	{
			var ag=document.frProducto.idAgte1.options[document.frProducto.idAgte1.selectedIndex].value;
			var so=document.frProducto.idsociedad1.options[document.frProducto.idsociedad1.selectedIndex].value;

			var x;
			var parametros;
			var num;

			if (so!=null)
			{
				if ( ag=='')
					{
						document.frProducto.nivel_detalle1.value='SO';
					}
				else
					{
						document.frProducto.nivel_detalle1.value='AG';
					}
		}

		document.frProducto.pagina0.value="./queryfilter.asp";
		document.frProducto.pagina1.value="./covercomision.asp";
	}
	
	document.frProducto.action=document.frProducto.pagina1.value;
	document.frProducto.submit();

	return true;
}

function ChangeidAgte1()
{
var selected ;

}

function ClickidSociedad1()
{
var selected;
}


function ChangeProducto1()
{
var ag=document.frProducto.Producto1.options[document.frProducto.Producto1.selectedIndex].value;
if (ag==''  ) {
		selected='';
		document.frProducto.nivel_PR.value=selected;
		}
	else
		{
		selected='PR';
		document.frProducto.nivel_PR.value=selected;
		}
}


function ChangeWsaler1()
{
var ag=document.frProducto.Wsaler1.options[document.frProducto.Wsaler1.selectedIndex].value;
if (ag==''  ) {
		selected='';
		document.frProducto.nivel_WS.value=selected;
		}
	else
		{
		selected='NC';
		document.frProducto.nivel_WS.value=selected;
		}
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
			Response.Write "<h3> Detalle Cover Page </h3>"
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
			
OpenTable "75%", "'' align=center"	
	OpenTr "class=teven"
		OpenTd "thead", "width=1% align=left"
		
			Response.Write "Periodo Consulta"
		CloseTd
	closetr
closetable

OpenTable "75%", "'' align=center"			
  OpenTr "valign=top"
	OpenTd "thead", "width=1% align=left"
		Response.Write "Fecha inicial "
	CloseTd
	
	OpenTd "thead", "width=1% align=left"
				Response.Write "Año:"
				OpenCombo "fromyear1",  "class=bttntext"
				%><!--		"<OPTION value=''> <% %>	<% %></OPTION>"--><%
					For I = 1999 To Year(Date())
						If I = year(date()) Then
							Sel="selected"
						Else
							Sel=""
						End If
								PlaceItem Sel, I, I
					Next
				CloseCombo
	closetd		
	OpenTd "thead", "width=1% align=left"				
				
				Response.Write "Mes:"
				OpenCombo "frommonth1",  "class=bttntext"  
				%><!--		"<OPTION value=''> <% %>	<% %></OPTION>"--><%
					For I = 1 To 12
						If I = month(date()) Then
							Sel="selected"
						Else
							Sel=""
						End If
								PlaceItem Sel, I, I
					Next
				CloseCombo
	CloseTd
	closetr
OpenTr "class=teven"

	OpenTd "thead", "width=1% align=left"
		Response.Write "Fecha final "
	CloseTd

		Set adoConn = GetConnpipelineDB()
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec Comision.dbo.Parametros_consultaPipelineComisiones_Select"

		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			fecFinalConsultaComision = Date()
		Else
			arrAgente = rstCampo.GetRows()
			fecFinalConsultaComision = arrAgente(1 , 0)
		End If
		
		rstCampo.Close
		adoConn.close
		Set rstCampo = nothing
		set adoconn=nothing

	
	OpenTd "thead", "width=1% align=left"
	
				Response.Write "Año:"
				OpenCombo "toyear1",  "class=bttntext"
				%><!--		"<OPTION value=''> <% %>	<% %></OPTION>"--><%
					For I = 1999 To Year(fecFinalConsultaComision)
						If I = year(date()) Then
							Sel="selected"
						Else
							Sel=""
						End If
								PlaceItem Sel, I, I
					Next
				CloseCombo
		closetd
		OpenTd "thead", "width=1% align=left"				
				Response.Write " Mes:"
				OpenCombo "tomonth1",  "class=bttntext"
				%><!--		"<OPTION value=''> <% %>	<% %></OPTION>"--><%
					For I = 1 To 12
						If I = month(date()) Then
							Sel="selected"
						Else
							Sel=""
						End If
								PlaceItem Sel, I, I
					Next
				
				CloseCombo
	closetd
			
closetr	
CLOSETABLE		
write_dataLog Response.Status,"queryfilter.asp", "queryfilter.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),Sql,"N/A","null","Consulta","N/A"

OpenTable "75%", "'' align=center"			

	%>
<!--		<SELECT NAME="Wsaler1"  class="bttntext" onChange="ChangeWsaler1();">
		<OPTION value="">-- No detallar por Wsaler  ---</OPTION>
		<OPTION value="ALL">-- Todos los Wsalers  ---</OPTION>
		--><%		
	'	For J = 0 To Ubound(arrWsaler, 2) 
	'		Response.Write "<OPTION value='" & trim( arrWsaler(0,J) )& "'> " & _
	'		arrWsaler(1,J)& "</OPTION>"
	'	Next    
		%>
		<!--
				</SELECT>-->
		<%
'	CloseTd
'
'CloseTr

OpenTr "valign=top"
	OpenTd "thead", "width=10% align=left"
		Response.Write "Sociedad" 
	CloseTd
	
	    
	OpenTd "", ""
	    OpenCombo "idsociedad1", " class=bttntext onclick='javascript:fillFP(this)' id='Sociedad'"
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
		OpenTd "thead", "width=10% align=left"
		Response.Write "Agente" 
		
	CloseTd
	OpenTd "tbody", "width=90% align=left"
		get_agente idagte, idsociedad, wsaler
		%><SELECT NAME="idAgte1"  class="bttntext" onChange="ChangeidAgte1();" ID="Select2">	<%		
			if todosag=1 then
				%>
					<OPTION selected value="">-- Cover Sociedad --</OPTION>
				<%		
		end if
		%>	</SELECT>	<%
	CloseTd

CloseTr

	
CLOSETABLE		

OpenTable "75%", "'' align=center height=14"
	OpenTr "class=thead"

		OpenTd "tbody2", "width=10% align=left"
			Response.Write "&nbsp;"
		CloseTd
		OpenTd "tbody2", "width=25% align=left"
				PlaceInput "Enviar ", "button", "Enviar ", "class=sbttn  onclick='go(" & Year(fecFinalConsultaComision) & " , " & Month(fecFinalConsultaComision) & ");'"
				PlaceInput "Regresar", "button", "Regresar","class=sbttn   onclick='back();'"
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
		
		PlaceInput "pagina1","hidden","../covercomision.asp",""'por defecto va a default
		PlaceInput "pagina0","hidden","./queryfilter.asp",""
		PlaceInput "pagina_2","hidden","../default.asp",""
	CloseTr


CloseForm
end if
CloseBody
CloseHTML


If Err.number <> 0 Then'
	Set adoConn = GetConnpipelineDB()

	'write_sp_log adoConn, 19101, "Error", 0, "", "", 0, 0, "", mid( "comisiones/queryfilter.asp Loaded by " & Session("sp_miLogin") & " err:" & err.Description ,1,250)
	write_sp_log adoConn, 13348, "Error", 0, "", "", 0, 0, "", mid( "comisiones/queryfilter.asp Loaded by " & Session("sp_miLogin") & " err:" & err.Description ,1,250)

	CloseConnPipelineDB

	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>