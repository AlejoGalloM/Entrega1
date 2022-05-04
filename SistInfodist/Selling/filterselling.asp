<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:					filterselling.asp  19306
'Path:						/sistinfodist/selling/filterselling.asp
'Created By:				Margarita Cardozo- Jimmy Ospino 2003/05/26
'Last Modified:				2003/07/30  --20038/08/26 mmc adicione log en on error
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
dim sql 'execute
dim sql1 'log sp name
dim sql3 'log
Dim adoConn,oConn'Database Connection
Dim rstcampo
Dim arrProductos, arrSociedades, arrAgente, arrWsaler
dim arribc, arrsal,arrSellingProductMix
Dim rs, cn,objRst,sel
Dim I,J,K,L      
Dim Idsociedad, idagte, Wsaler
dim idwsaler
Dim Page 

dim nc
dim parametros
dim idsociedad1,idagte1,tengaproductos1,notengaproductos1,productorango1,rango1,ordenepor1

dim Wsaler1,option1
dim todosag, todosso
dim nrodocumemp1,tipodocumemp1,nombreemp1

dim conexion

Authorize 1,25
Set adoConn = GetConnPipelineDB
write_sp_log adoConn, 19306, "", 0, "", "", 0, 0, "", "selling/filterselling.asp Loaded by " & Session("sp_miLogin") & " cargando pagina"
CloseConnPipelineDB
Set adoConn = Nothing
'======================================================================================
'parameters
'======================================================================================
nrodocumemp1=		Request.Form ("nrodocumemp1")
tipodocumemp1=		Request.Form ("tipodocumemp1")
nombreemp1=			Request.Form ("nombreemp1")

tengaproductos1=	Request.Form ("tengaproductos1")
notengaproductos1=	Request.Form ("notengaproductos1")
productorango1=		Request.Form ("productorango1")
rango1=				Request.Form ("rango1")
ordenepor1=			Request.Form ("ordenepor1")
option1			=	Request.Form ("option1")
idagte1=			Request.Form ("idagte1")
idsociedad1	=		Request.Form ("idsociedad1")

if isnull(Wsaler1) then
	  	Wsaler1="0"
end if

if isnull(nrodocumemp1) or len(nrodocumemp1)=0 then
	  	nrodocumemp1=0
end if
   		
if isnull(tipodocumemp1) then
	  	tipodocumemp1=""
end if

if isnull(idsociedad1) then
	  	idsociedad1=0
end if

if isnull(idagte1) then
	  	idagte1=0
end if

if isnull(tengaproductos1) then
	  	tengaproductos1=""
end if

if isnull(notengaproductos1) then
	  	notengaproductos1=""
end if

if isnull(productorango1) then
	  	productorango1=""
end if

if isnull(rango1) then
	  	rango1=0
end if

if isnull(ordenepor1) then
	  	ordenepor1=""
end if

'======================================================================================
'security
'======================================================================================

dim AccessLevel, idagteLoggedIn, idSociedadLoggedIn, idagteContract, idSocContract 
dim isAuthorized 
AccessLevel= Cstr(Session("sp_AccessLevel"))
idagteLoggedIn= CStr(Session("sp_idagte"))
idSociedadLoggedIn= CStr(Session("sp_Idsoc"))
Wsaler=""
idagte=idagteLoggedIn
idsociedad=idSociedadLoggedIn


isAuthorized  = false
Select Case AccessLevel
		Case 0 'Skandia
				Wsaler=null
				idagte=0
				idsociedad=0
				todosso=1
				todosag=1
				isAuthorized  = true		
		Case 1 'WHOLE SALER	
				Wsaler = get_WsalerWsaler(Session("sp_Idworker"))
				
				if not isnull(Wsaler) and wsaler<>"" and wsaler <>"0" then
					isAuthorized = true			
				end if	
				todosso=1
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
		response.end
		Response.Redirect Application("UnauthorizedURL")		  
end if
'-------------------------------------------------------------
get_agente  idagte, idsociedad,Wsaler
get_sociedad idsociedad,Wsaler
get_SellingProductMixget()
get_IBC()
get_SAL()
get_producto()

write_dataLog  Response.Status,"filterselling.asp", "filterselling.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),conexion,"N/A","null","Consulta","N/A"


function get_WsalerWsaler( idworker )
		dim rstcampo1
		dim adoconn1
		dim arrWsaler
		Set adoConn1 = GetConnpipelineDB()		
		Set rstCampo1 = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.WsalerGetWsaler " & idworker
		sql3= replace(Sql,"'","''")
		sql3= mid( "sistinfodist/selling/filterselling.asp Loaded by " & Session("sp_miLogin") & "- par: " & sql3 ,1,250)
		write_sp_log adoConn1, 19306, "sigscg.dbo.WsalerGetWsaler", 0, "", "", 0, 0, "",sql3	
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
		conexion=conexion&" - "&Sql
		
end function


function get_sociedad( idsociedad,Wsaler )
		Set adoConn = GetConnpipelineDB()		
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.spcm_getsociedadfp " & idsociedad & ",'" & Wsaler & "'"
		sql3= replace(Sql,"'","''")
		sql3= mid( "sistinfodist/selling/filterselling.asp Loaded by " & Session("sp_miLogin") & "- par: " & sql3 ,1,250)
		write_sp_log adoConn, 19306, "sigscg.dbo.spcm_getsociedadfp", 0, "", "", 0, 0, "",  sql3
		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			arrSociedades = 0
		Else
			arrSociedades = rstCampo.GetRows()
		End If
		conexion=conexion&" - "&Sql
		rstCampo.Close
		adoConn.close
		Set rstCampo = nothing
		set adoconn=nothing
end function


function get_agente(idagte, idsociedad,Wsaler )
		Set adoConn = GetConnpipelineDB()		
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.spcm_getagente " & idagte & "," & idsociedad & ",'" & Wsaler & "'"
		sql3= replace(Sql,"'","''")
		sql3= mid( "sistinfodist/selling/filterselling.asp Loaded by " & Session("sp_miLogin") & "- par: " & sql3 ,1,250)
		write_sp_log adoConn, 19306, "sigscg.dbo.spcm_getagente", 0, "", "", 0, 0, "",  sql3
		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			arrAgente = 0
		Else
			arrAgente = rstCampo.GetRows()
		End If
		rstCampo.Close
		adoConn.close
		conexion=Sql
		Set rstCampo = nothing
		set adoconn=nothing
end function

function get_producto()
		Set adoConn = GetConnpipelineDB()		
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.spcm_getproducto " 
		conexion=conexion&" - "&Sql
		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			
			arrProductos = 0
		Else
			arrProductos = rstCampo.GetRows()
		End If
		rstCampo.Close
		adoConn.close
	
		Set rstCampo = nothing
		set adoconn=nothing
end function


function get_IBC()
		Set adoConn = GetConnpipelineDB()		
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.spcm_rangosaldo_get 'IBC' "
		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			arrIBC = 0
		Else
			arrIBC = rstCampo.GetRows()
		End If
		rstCampo.Close
		adoConn.close
		conexion=conexion&" - "&Sql
		Set rstCampo = nothing
		set adoconn=nothing
end function

function get_SAL()
		Set adoConn = GetConnpipelineDB()		
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.spcm_rangosaldo_get 'SAL' "
		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			arrSAL = 0
		Else
			arrSAL = rstCampo.GetRows()
		End If
		rstCampo.Close
		adoConn.close
		conexion=conexion&" - "&Sql
		Set rstCampo = nothing
		set adoconn=nothing
end function


function  get_SellingProductMixget()
		Set adoConn = GetConnpipelineDB()		
		Set rstCampo = Server.CreateObject("ADODB.Recordset")
		Sql = "exec sigscg.dbo.SellingProductMix_get"
		rstCampo.Open Sql, adoConn
		If rstCampo.BOF And rstCampo.EOF Then
			arrSellingProductMix = 0
		Else
			arrSellingProductMix = rstCampo.GetRows()
			
		End If
		rstCampo.Close
		adoConn.close
		conexion=conexion&" - "&Sql
		Set rstCampo = nothing
		set adoconn=nothing
		
end function


sub crossparameters()
	
	OpenTable "80%", "'' align=center"	
		OpenTr "class=teven valign=top"
			OpenTd "thead", "width=30% align=left"
				Response.Write "Productos a cruzar" 
			CloseTd
			OpenTd "", "width=70%"
			    OpenCombo "sellingproductmix1", " class=bttntext onclick='javascript:fillMix(this)'"
			    	sel=""
					J = 0	    
					if  UBound(arrSellingProductMix, 2)>=0 then
						PlaceItem "selected", arrSellingProductMix(0,J),arrSellingProductMix(0,J)
						For J = 1 To UBound(arrSellingProductMix, 2)
							PlaceItem Sel, arrSellingProductMix(0,J),arrSellingProductMix(0,J)
						Next 'J
					end if		
				CloseCombo
			CloseTd
		closetr
	closetable
	OpenTr "class=teven"'
		OpenTd "thead", " align=left"
			Response.Write "&nbsp;"
		CloseTd
	closetr
	OpenTr "class=teven"'
		OpenTd "thead", " align=left"
			Response.Write "&nbsp;"
		CloseTd
	closetr
	
	OpenTable "80%", "'' align=center"	
		OpenTr ""'
			OpenTd "thead", " align=left"
				Response.Write "CONDICIONALES (Selling)" 
			CloseTd
		closetr
		OpenTr ""'
			OpenTd "thead", " align=left"
				Response.Write "&nbsp;"
			CloseTd
		closetr
		OpenTr ""'
			OpenTd "tbody", "align=left"
				Response.Write "<input type=Radio Button name=selling onClick='activartiposaldo()' checked>"
				Response.Write "Tenga todos los productos seleccionados (Up-selling)"			
			CloseTd
		closetr
	closetable				
	OpenTable "50%", "'' align=center"	
	   OpenTr ""'
			OpenTd "tbody", ""
				Response.Write "Ordenar por y seleccionar el rango para  " 
			CloseTd
			OpenTd "",""
			    OpenCombo "Orden1", " class=bttntext onclick='javascript:activartiposaldo()'"
			   		sel=""
						PlaceItem "selected", "1","Primer producto"
						PlaceItem Sel, "2","Segundo producto"
				CloseCombo
			CloseTd
		closetr
	closetable				
	
	OpenTable "80%", "'' align=center"	
		OpenTr "class=todd"
			OpenTd "tbody", " align=left"
				Response.Write "<input type=Radio Button name=selling onClick='activartiposaldo()'>"
				Response.Write "Tenga solamente el primer producto de la selección y no el segundo (Cross-selling)"			
			CloseTd
		closetr
		OpenTr ""'
			OpenTd "tbody", "width=27% align=left"
				Response.Write "<input type=Radio Button name=selling onClick='activartiposaldo()' >"
				Response.Write "Tenga solamente el segundo producto de la selección y no el primero (Cross-selling)"			
			CloseTd
		CloseTr
		OpenTr "class=teven"'
			OpenTd "thead", " align=left"
				Response.Write "&nbsp;"
			CloseTd
		closetr
	closetable				

	OpenTable "80%", "'' align=center"
		OpenTr "valign=top"
			OpenTd "thead", "width=40% align=left"
				Response.Write("<br>")
				Response.Write "SELECCIONAR SEGUN RANGO ** " 
			CloseTd
		closetr
		OpenTr ""'
			OpenTd "thead", " align=left"
				Response.Write "&nbsp;"
			CloseTd
		closetr
		OpenTr "class=todd"'
			OpenTd "", " align=left"
				Response.Write "<input type=Radio Button   name=OpRango onClick='activartiposaldo()' checked >" 
				Response.Write ("Rango IBC")
			CloseTd
		    
			OpenTd "", " align=left"
			    OpenCombo "RangoIBC1", ""
			    	sel=""
			    	J=0
			    	PlaceItem "selected", 0 ," Cualquier rango "
			    	if  UBound(arrIBC, 2)>=0 then
						PlaceItem "selected", arrIBC(1,J) ,arrIBC(1,J) & ". " & arrIBC(2,J) 
					end if		
					For J = 1 To UBound(arrIBC, 2)
						PlaceItem Sel, arrIBC(1,J) ,arrIBC(1,J)& ". " & arrIBC(2,J) 
					Next 'J
				CloseCombo
			CloseTd
		closetr
		OpenTr "class=todd"'
			OpenTd "", " align=left"
				Response.Write "<input type=Radio Button  name=OpRango onClick='activartiposaldo() ' >" 
				Response.Write "Rango Saldo" 
			CloseTd
			
			OpenTd "", " align=left"
				 OpenCombo "RangoSAL1", ""
					sel=""
					J=0
					PlaceItem "selected", 0 ," Cualquier rango "
					For J =0 To UBound(arrSAL, 2)
						 if rango1=arrSAL(1,J) then  'hay valor defecto y es igual a este item
							PlaceItem "selected", arrSAL(1,J) ,arrSAL(1,J)  & ". " &  arrSAL(2,J) 
						 else
							PlaceItem "", arrSAL(1,J) ,arrSAL(1,J)  & ". " &  arrSAL(2,J) 
						 end if	
					Next 'J
				CloseCombo
			CloseTd
					
		closetr
		OpenTr "class=teven"'
			OpenTd "thead", " align=left colspan=2"
				Response.Write "&nbsp;"
			CloseTd
		closetr
			
	closetable

	OpenTable "80%", "'' align=center"
		OpenTr "valign=top"
			OpenTd "tbody2", ""
				Response.Write "*El listado se ordenará de mayor a menor y según la condición."
			CloseTd
		OpenTr "valign=top"
	closetr		    
	
	OpenTd "tbody2", ""
		CloseTd
		closetr
		closetable

		OpenTable "80%", "'' align=center"
			OpenTr "valign=top"
				OpenTd "tbody2", ""
					Response.Write "** Si seleccionó Fondo Obligatorio el reporte estará ordenado por IBC - Ingreso Base de cotización"
				CloseTd

'			OpenTr "valign=top"
'			closetr		    
'				OpenTd "",""
'				CloseTd
			closetr
		closetable
end sub

sub ag_so

	OpenTable "80%", "'' align=center"			
		OpenTr "class=teven"'
			OpenTd "thead", "width=30% align=left"
				Response.Write "Sociedad" 
			CloseTd
			
			OpenTd "", "width=70%"
				OpenCombo "idsociedad0", " class=bttntext onclick='javascript:fillFP(this)' id='Sociedad'"
		   			if todosso=1 then
   		  				PlaceItem "", 0,"-- Todas las sociedades --"	    
					end if
					For J = 0 To UBound(arrSociedades, 2)
						If CStr(Request.Form("idsociedad1")) = CStr(arrSociedades(0,J)) Then
							Sel = "selected"
						Else
							Sel = ""
						End If
						PlaceItem Sel, arrSociedades(0,J), arrSociedades(1,J)
					Next 'J
				CloseCombo
			CloseTd
		CloseTr

		OpenTr "class=todd"'
			OpenTd "thead", "width=30% align=left"
				Response.Write "Agente" 
			CloseTd
			
			OpenTd "tbody", "width=70% align=left"
				get_agente idagte, idsociedad, Wsaler
				%><SELECT NAME="idagte0" class="bttntext" onChange="Changeidagte0();" ID="Select2">	<%		
	   				if todosag=1 then
						%>
						<OPTION selected value="0">-- Todos los agentes --</OPTION>
						<%		
					end if
				%></SELECT><%
			CloseTd
		CloseTr
	closetable

end sub

sub empresa
	OpenTable "80%", "'' align=center"			
	if isnull( nrodocumemp1) or nrodocumemp1<1 then
		OpenTr ""
			OpenTd "thead", "width=100% align=rigth"
		 		Response.Write "La empresa que seleccionó no es válida"
			CloseTd
		closetr	
		'OpenTr ""
	else
		OpenTr "class=teven"'
			OpenTd "thead", "width=30% align=left"
				Response.Write "Empresa" 
			CloseTd
			OpenTd "thead", "width=70% align=rigth"
		 		Response.Write  nombreemp1  & vbCrLf
			CloseTd
		closetr	
		OpenTr ""
			OpenTd "thead", " align=left"
				Response.Write "NIT" 
			CloseTd
			OpenTd "thead", " align=left"
				Response.Write  nrodocumemp1 & "  " & tipodocumemp1  & vbCrLf 
			CloseTd
		closetr
	end if
	Closetable
end sub


OpenHTML

OpenHead

PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
PlaceMeta "Pragma", "", "no_cache"

%>

<script language="javascript" src="/super_pipeline/operations/_pipeline_scripts/SkCoSecurity.js"></script>

<%

%>
<SCRIPT LANGUAGE="javascript">

//'array sociedades lado cliente
//var Sociedades = new Array(<%=UBound(arrSociedades, 2)+2 %>)  //dim tamaño de filas

var largo=<%=UBound(arrSociedades, 2) + 2 %>
var Sociedades = new Array(largo)  //dim tamaño de filas
//aqui voy a adicionar todas las sociedades
for (i = 0; i < largo; i ++) {  //por eso arranca en 1
	Sociedades[i] = new Array(<%=UBound(arrSociedades,1)%>);
}
<%

if todosso=1 then
	Response.Write"Sociedades[0][0] = ' '" & vbCrLf
	Response.Write"Sociedades[0][1] = 'Todas las sociedades'" & vbCrLf
	Response.Write"Sociedades[0][2] = ''" & vbCrLf
	Response.Write"Sociedades[0][3] = ''" & vbCrLf
end if
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
var agentes = new Array(largo)
for (i = 0; i < largo; i ++) {
		agentes[i] = new Array(<%=UBound(arrAgente,1)%>);
	}
<%
if todosag=1 then
		Response.Write"agentes[0][0] = ' '" & vbCrLf
		Response.Write"agentes[0][1] = 'Todos los agentes'" & vbCrLf
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
	    document.frProducto.idagte0.length = 1
	  //  alert (String(agentes.length));
		for (i = 0; i < agentes.length; i ++) 
		{ 
			if (agentes[i][2] == Sociedades[Fp][1]) 
				{   
				     addoption(agentes[i][0], agentes[i][1],document.frProducto.idagte0)
				}
		}
	return true;
}


function fillMix(obj) 
{
	var Fp = obj.selectedIndex;
	
	if (document.frProducto.option1.value==2 || document.frProducto.option1.value==1)
	{
		activartiposaldo();
	}
	return true;
}




function getRadioValue(radioName) {
		var collection;
		var j = -1;
			
		collection = document.frProducto.elements;

		for (i=0;i<collection.length;i++) {
			if (collection[i].type == "radio" && collection[i].name == radioName){
			j = j + 1;
			if (collection[i].checked)
				return(j);
			}	
		}
		return j;
	}



function ActivateRadioValue(radioName) {
		var collection;
		var j = -1;
			
		collection = document.frProducto.elements;

		for (i=0;i<collection.length;i++) {
			if (collection[i].type == "radio" && collection[i].name == radioName){
			j = j + 1;
			collection[i].enabled=false;
			}	
		}
		return j;
	}

	
function activartiposaldo() 
{	
	var ss;
	var rango;
	var len,i;
	var fpob;
	var buscar=0;
	var tengaproducto=0; 
	var orden=0;
	
	if ( (document.frProducto.idagte0.selectedIndex==null ) )
	{
			document.frProducto.idagte0.selectedIndex==0
	}
		
	if (document.frProducto.productorango1.value==null)
		{
		 	document.frProducto.productorango1.value=""
		}


	if (document.frProducto.option1.value==2 || document.frProducto.option1.value==3)
	{
    document.frProducto.OpRango.disabled = true;
    ActivateRadioValue('OpRango');
          
	var selected1 = getRadioValue('selling');
	var selected2 = getRadioValue('OpRango');
	var so=document.frProducto.sellingproductmix1.options[document.frProducto.sellingproductmix1.selectedIndex].value;
	//en ss tengo el arreglo de los productos que selecciono para armar el where
	ss = so.split("-");
	len=ss.length;
	
	if (document.frProducto.ordenepor1.value==null)
		{
		 	document.frProducto.ordenepor1.value=""
		}
	
	//cual producto quiere filtrar
	
	if (selected1<0)
	{
			alert ("Por favor seleccione un condicional");
			document.selling.focus();
			return;
	}
	
	if (document.frProducto.Orden1.value==1) 
			{ orden=0
			
			}
	else
			{	orden=1
			
			}
	
	switch (selected1)
	 {
	 
	  case 0:
			document.frProducto.notengaproductos1.value="" ;
		  	tengaproducto=0;//tenga todos los productos
  			document.frProducto.tengaproductos1.value=so  ;
  			document.frProducto.productorango1.value=ss[orden];
			document.frProducto.ordenepor1.value=ss[orden];
   			buscar=0; 
   			document.frProducto.Orden1.disabled = false;
 			document.frProducto.Orden1.value=document.frProducto.Orden1.options[document.frProducto.Orden1.selectedIndex].value;
			document.frProducto.ordenepor1.value=ss[orden];
			document.frProducto.productorango1.value=ss[orden];
		
   			if (ss[orden]=='FPOB')//si el primero es obligatorio
					{
						 document.frProducto.OpRango[0].disabled = false;
						 document.frProducto.OpRango[0].checked =  true;
						 
						 document.frProducto.OpRango[1].disabled = true;
						 
						 document.frProducto.RangoIBC1.disabled = false;
						 document.frProducto.RangoSAL1.disabled = true;
						 document.frProducto.rango1.value=document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].value;
						 document.frProducto.descrango1.value 	= document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].text; 
					}
			else 
					{
						 document.frProducto.OpRango[1].disabled = false;
						 document.frProducto.OpRango[1].checked =  true;
						 
						 document.frProducto.OpRango[0].disabled =  true;
						 
						 document.frProducto.RangoIBC1.disabled = true;
						 document.frProducto.RangoSAL1.disabled = false;
						 document.frProducto.rango1.value=document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].value;
						 document.frProducto.descrango1.value 	= document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].text; 
					}		
			break
	 
  	 case 1:
			document.frProducto.Orden1.disabled = true;
			document.frProducto.tengaproductos1.value=ss[0]  ;
   		  	document.frProducto.notengaproductos1.value=ss[1]  ;
   		  	document.frProducto.productorango1.value=ss[0];
			document.frProducto.ordenepor1.value=ss[0];
   	
   		  	buscar=1; //tenga solo el primer producto
			if (ss[0]=='FPOB')//si el primero es obligatorio
					{
						 document.frProducto.OpRango[0].disabled = false;
						 document.frProducto.OpRango[0].checked =  true;
						 
						 document.frProducto.OpRango[1].disabled =  true;
	 					 
	 					 document.frProducto.RangoIBC1.disabled = false;
						 document.frProducto.RangoSAL1.disabled = true;
						 
						 document.frProducto.rango1.value=document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].value;
						 document.frProducto.descrango1.value 	= document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].text; 
					}
			else 
					{
						 document.frProducto.OpRango[1].disabled = false;
						 document.frProducto.OpRango[1].checked =  true;
						 
						 document.frProducto.OpRango[0].disabled =  true;
						 
						 document.frProducto.RangoIBC1.disabled = true;
						 document.frProducto.RangoSAL1.disabled = false;

						 document.frProducto.rango1.value=document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].value;
						 document.frProducto.descrango1.value 	= document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].text; 
					}		
			break
		 
     case 2:
			document.frProducto.Orden1.disabled = true;
			document.frProducto.tengaproductos1.value=ss[1];
		  	document.frProducto.notengaproductos1.value=ss[0]  ;
		  	document.frProducto.productorango1.value=ss[1];
			document.frProducto.ordenepor1.value=ss[1];
   			tengaproducto=2;//tenga 2
   			buscar=2; //tenga solo el segundo producto
			if (ss[1]=='FPOB')//el segundo es obligatorio
					{
						 document.frProducto.OpRango[0].disabled = false;
						 document.frProducto.OpRango[0].checked =  true;
						 
						 document.frProducto.OpRango[1].disabled =  true;
						 
						 document.frProducto.RangoIBC1.disabled = false;
						 document.frProducto.RangoSAL1.disabled = true;

						 document.frProducto.rango1.value=document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].value;
						 document.frProducto.descrango1.value 	= document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].text; 
					}
			else 
					{
						 document.frProducto.OpRango[1].disabled = false;
						 document.frProducto.OpRango[1].checked =  true;
						 
						 document.frProducto.OpRango[0].disabled = true;
						 	 
						 document.frProducto.RangoIBC1.disabled = true;
						 document.frProducto.RangoSAL1.disabled = false;

						 document.frProducto.rango1.value=document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].value;
						document.frProducto.descrango1.value 	= document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].text; 
					}		
			break
		
	 default:
		{	 
  			buscar=0;//eligio  tenga stodos los producto sellecionados hay que ver si el primero es fpob
  			
  			if (ss[0]=='FPOB')//si el primero es obligatorio
				{
					 document.frProducto.OpRango[0].checked = true;
					 document.frProducto.OpRango[1].disabled = false;
					 document.frProducto.rango1.value=document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].value;
			  		document.frProducto.descrango1.value 	= document.frProducto.RangoIBC1.options[document.frProducto.RangoIBC1.selectedIndex].text; 
				}
			else 
				{
					 document.frProducto.OpRango[0].checked = false;
					 document.frProducto.OpRango[1].disabled = true;
					 document.frProducto.rango1.value=document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].value;
					 document.frProducto.descrango1.value 	= document.frProducto.RangoSAL1.options[document.frProducto.RangoSAL1.selectedIndex].text; 
				}		
			break
		 }
     }
  }
  else
  {
	document.frProducto.ordenepor1.value= document.frProducto.productoorder.options[document.frProducto.productoorder.selectedIndex].value;
  }
	document.frProducto.idsociedad1.value= document.frProducto.idsociedad0.options[document.frProducto.idsociedad0.selectedIndex].value;	
	document.frProducto.nombresoc1.value= document.frProducto.idsociedad0.options[document.frProducto.idsociedad0.selectedIndex].text;	
	document.frProducto.idagte1.value = document.frProducto.idagte0.options[document.frProducto.idagte0.selectedIndex].value;	
    document.frProducto.nombreagte1.value 	= document.frProducto.idagte0.options[document.frProducto.idagte0.selectedIndex].text;
	return true;
}


function back()
{

	if (document.frProducto.option1.value==2 || document.frProducto.option1.value==1)
	{
			frProducto.action="./typeofquery.asp";
			
	}
	else
	{
			frProducto.action="./default.asp";
	}
	document.frProducto.submit();
	
}

function go()
{
	var selected ;
	
	
	if ( (document.frProducto.idagte0.selectedIndex==null )  )
	{		
			alert ("Por favor elija un agente válido");
			document.frProducto.idagte0.focus();
			frProducto.action="./filterselling.asp";
			return;
	}
	else
	{
			if  (document.frProducto.idagte0.selectedIndex==-1) 
			{
				alert ("Por favor elija un agente válido");
				document.frProducto.idagte0.focus();
				frProducto.action="./filterselling.asp";
				return;
			}
		
	
  	 }
		
	activartiposaldo();	
	document.frProducto.pagina0.value="./filterselling.asp";
	document.frProducto.pagina1.value="./covercomision.asp";
	frProducto.action="./crossselling.asp";
	//document.frProducto.submit();
	
	return true;
}

function Changeidagte0()
{
var selected ;

}

function Clickidsociedad0()
{  
var selected
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

</SCRIPT>
<%
PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
PlaceLink "REL", "stylesheet", "../../css/style1.css", "text/css"
closehead

OpenTable "80%", "'' align=center cellpadding=0 cellspacing=0"
	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
			Response.Write "&nbsp;"
		CloseTd
	CloseTr
	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=2"
			Response.Write "<h3>Consulta"
			Select case option1
				Case "1"  response.Write " Empresa Clientes "
				Case "2"  response.Write " Cruzada por  Empresa "
				Case "3"  response.Write " Cruzada por Clientes Individuales "
				Case Else      
					response.Write "Usuario no autorizado"
			End Select
			Response.Write "</h3>"
		CloseTd
	CloseTr
	
	OpenTr "valign=middle"
		OpenTd "''", "align=center valign=middle colspan=1"
			Response.Write "<hr>"
		CloseTd
	CloseTr
closetable	

OpenTable "80%", "'' align=center border=0 "
	OpenTr "class=tbody2"
		OpenTd "", "width=20% align=right"
			Response.Write "Usuario: " 									
			Response.Write (Session("sp_Usuario"))
		closetd
	closetr
closetable

OpenForm "frProducto", "post", "./default.asp", ""

If autorizarMn(1,25) Then

	Select Case option1
		Case "1" 
			empresa
			ag_so
			OpenTable "80%", "'' align=center"	
				OpenTr "valign=top class=teven"
					OpenTd "thead", "width=30% align=left"
						Response.Write "Ordenar por Producto " 
					CloseTd
					OpenTd "thead", "width=70% align=left"
						OpenCombo "productoorder", " class=bttntext onclick='javascript:fillMix(this)'"
		   					sel=""
							j=0	    
							For J = 0 To UBound(arrProductos, 2)
								PlaceItem Sel, arrProductos(0,J),arrProductos(0,J)
							Next 'J
						CloseCombo
					CloseTd
				closetr
			closetable
			isAuthorized=true
		Case "2" 
			empresa
			ag_so
			crossparameters
			isAuthorized=true
		case "3"
			ag_so
			crossparameters
			isAuthorized=true
		case else
			isAuthorized=false
	End Select

	If isAuthorized=false Then
			Response.Write("Acceso no autorizado")	
			Response.Redirect Application("UnauthorizedURL")		  
	end if

end if
closetable
	OpenTr "class=thead"
		OpenTd "tbody2", "width=30% align=left"
			Response.Write "&nbsp;"
		CloseTd
	closetr
	OpenTr "class=thead"
		OpenTd "tbody2", "width=30% align=left"
				Response.Write "&nbsp;"
		CloseTd
	closetr
	OpenTable "80%", "'' align=center height=14"

	OpenTr "class=thead"
		OpenTd "tbody2", "width=30% align=left"
			Response.Write "&nbsp;"
		CloseTd
	
		OpenTd "tbody2", "width=20% align=left"
				PlaceInput "Buscar", "submit", "  Buscar  ", "class=sbttn  onclick=go()"
		CloseTd

		OpenTd "tbody2", "width=20% align=left"
				PlaceInput "Volver", "submit", "  Volver  ","class=sbttn   onclick=back();"
		CloseTd

		OpenTd "tbody2", "width=30% align=left"
			Response.Write "&nbsp;"
		CloseTd
	CloseTr

	OpenTr "class=todd"
		PlaceInput "nombreemp1","hidden",Request.Form ("nombreemp1"),""
		PlaceInput "nombresoc1","hidden",Request.Form ("nombresoc1"),""
		PlaceInput "nombreagte1","hidden",Request.Form ("nombreagte1"),""
		PlaceInput "descrango1","hidden",Request.Form ("descrango1"),""
		PlaceInput "nrodocumemp1","hidden",Request.Form ("nrodocumemp1"),""
		PlaceInput "tipodocumemp1", "hidden", Request.Form ("tipodocumemp1"), ""
		PlaceInput "Wsaler1","hidden",Wsaler,Wsaler1				
		PlaceInput "idsociedad1","hidden",Request.Form ("idsociedad1"),""
		PlaceInput "idagte1","hidden",Request.Form ("idagte1"),""
		PlaceInput "tengaproductos1","hidden",Request.Form ("tengaproductos1"),""
		PlaceInput "notengaproductos1","hidden",Request.Form ("notengaproductos1"),""
		PlaceInput "productorango1","hidden",Request.Form ("productorango1"),""
		PlaceInput "rango1","hidden",Request.Form ("rango1"),""
			
		PlaceInput "ordenepor1","hidden",Request.Form ("ordenepor1"),""
		PlaceInput "tengaproductosexc1","hidden",Request.Form ("tengaproductosexc1"),""
		
		PlaceInput "pagina1","hidden","../covercomision.asp",""'por defecto va a default
		PlaceInput "pagina0","hidden","./filterselling.asp",""
		PlaceInput "pagina_2","hidden","../default.asp",""
		PlaceInput "option1","hidden",request.Form("option1"),""
	CloseTr

CloseForm

CloseBody
CloseHTML


If Err.number <> 0 Then'
	Set bc = Server.CreateObject("MSWC.BrowserType")
	Set adoConn = GetConnpipelineDB()
	write_sp_log adoConn, 19306, "Error", 0, "", "", 0, 0, "", mid("selling/filterselling.asp Loaded by " & Session("sp_miLogin") & " err:" & err.Description ,1,250)
	CloseConnPipelineDB
	
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
