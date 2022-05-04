<%@ Language=VBScript %>
<%
'===================================================================================
'@author name:		 		J carreno 
'@exception name:			include files does not exist	
'@param name description:	contract, product, plan, clientid, name	, listbeneficiary
'@return					list of beneficiarys of insurance
'@since						2002/04/17
'@version					1.0
'@File Name:				infoInsurance.asp [12900]
'@Path:						insurance/info
'revision					Julio 26 2002
'@Modified					J Carreño, cajulio@skandia.com.co, add document type 2003/08/26
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

Authorize 1,17

dim info(21,2)   'request 
dim infoseguro
dim ocultos
dim miSql 'SQL Sentences holder
dim rs
dim cn
dim resumen
dim cadenaComboDocumento
Dim pcmd			' ADODB.Command Object

sub flowHtml (mensaje)
OpenHTML
	OpenHead					  
		PlaceLink "REL", "stylesheet", "../../css/style1.css", "text/css"
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	closehead		

	OpenTable "''", "'' align=center"
			OpenTr ""
				OpenTd "'thead'", ""
				  Response.Write "<br><br>" & mensaje
				CloseTd
			CloseTr
	CloseTable
	
closehtml
end sub

sub cerrar (mensaje)
  flowHtml mensaje
  CloseConnpipelineDB     
	if rs.State=1 then
		rs.Close
	end if
	set rs=nothing
	Response.End
end sub

'===================================================================================
'==functions
'===================================================================================
sub escriba (mensaje)
  response.Write mensaje & "<br>"
end sub

function age (fecha)
   Dim intTemp 

   intTemp = DateDiff("yyyy", fecha, Date)

   If (Date < DateSerial(Year(Date), Month(fecha), Day(fecha))) Then
    intTemp = intTemp - 1
   End If

   age = intTemp
end function  

function NameDocumenType(doctype)
Dim DescripDocum 
Dim pos
dim pos2


	cadenaComboDocumento = PlaceDocTypeCombo("''", cn, "")
	cadenaComboDocumento = replace(cadenaComboDocumento, "DocType", "BenDocType") 'cambiar el nombre
	cadenaComboDocumento = replace(cadenaComboDocumento, "OPTION  value", "OPTION value") 'quitar los 2 espacios entre OPTION  value
	cadenaComboDocumento = replace(cadenaComboDocumento, "'", chr(34))
	pos = InStr(1, cadenaComboDocumento, "value="  & chr(34) & doctype & chr(34))
	pos2 = instr(pos, cadenaComboDocumento, "</OPTION>")
	if pos > 0 then
	    if pos2 > pos then
			DescripDocum = Mid(cadenaComboDocumento, pos + 10, pos2 - (pos + 10)) 'pos + 10 desde value='C' hasta > 
		else
			DescripDocum = "Indefinido"
		end if
	else
		DescripDocum = "Indefinido"
	end if
	NameDocumenType = DescripDocum
end function
'===================================================================================
'==get parameters
'===================================================================================
info(1,2)=Request.Form("Name")
info(2,2)=Request.Form("ClientId")
info(3,2)=Request.Form("Product")
info(4,2)=Request.Form("Contract")
info(5,2)=Request.Form("Plan")
'===================================================================================
'==set objects
'===================================================================================
Set cn = GetConnpipelineDB
Set rs = Server.CreateObject("ADODB.Recordset")
Set pcmd = server.CreateObject("ADODB.Command")
'===================================================================================
'==write log
'===================================================================================
write_sp_log cn, 12900, "", info(4,2), info(3,2), info(5,2), info(2,2), 0, "", "infoInsurance.asp loaded by " & Session.Contents("sp_milogin")

write_dataLog Response.Status,"infoInsurance.asp","infoInsurance.asp loaded by " & Session.Contents("sp_milogin"), Session.contents("name"),"Insurance..Insurance_Select - Insurance..sp_insu_adm_stateInsurance - Insurance..sp_insu_adm_amparo - Insurance..Beneficiary_Select", "N/A", "null", "Consulta","-"

'===================================================================================
'==process info insurance 
'===================================================================================
'misql="exec Insurance..Insurance_Select @producto='"&info(3,2)&"',@contrato="&info(4,2)&",@planproducto='"&info(5,2)&"'"
pcmd.CommandText = "Insurance..Insurance_Select"
pcmd.CommandType = 4 'adCmdStoredProc
pcmd.ActiveConnection = cn
pcmd.Parameters("@producto") = pcmd.CreateParameter("@producto", 200, 1, 7, info(3,2)) '200=adVarChar, 1=adInput
pcmd.Parameters("@contrato") = pcmd.CreateParameter("@contrato", 131, 1,18, info(4,2)) '131=adNumeric, 1=adInput
pcmd.Parameters("@contrato").Precision=0
pcmd.Parameters("@contrato").NumericScale=18
pcmd.Parameters("@planproducto") = pcmd.CreateParameter("@planproducto", 200, 1, 7, info(5,2)) '200=adVarChar, 1=adInput
set rs = pcmd.Execute
'rs.Open misql, cn

if rs.EOF and rs.BOF then
  cerrar "No tiene producto Asegurado"
end if
'===================================================================================
'==fill parameters
'===================================================================================
info(6,2) = Rs("nombre")
info(7,2) = Rs("identificacion")
info(8,2) = Rs("fechanac")
info(9,2) = age(info(8,2)) 
'-------------------------------------------
info(10,2) = Rs("metaahorro")
info(11,2) = Rs("valorasegurado")
info(12,2) = Rs("codamparo")
info(13,2) = Rs("valorprima")  
info(14,2) = Rs("valorprimamensual")  
info(15,2) = Rs("fechainicio")
info(16,2) = Rs("estadoSeguro")

'2003/08/26 J carreño
info(18,2) = Request.Form("DocType")
info(19,2) = NameDocumenType(info(18,2))
'primero la funcion que construye la variable cadenacombodocumento
info(21,2) = NameDocumenType(Rs("tipodocum"))
info(20,2) = ""
'fin 2003/08/26

'===================================================================================
'==get info state
'===================================================================================
'misql="Insurance..sp_insu_adm_stateInsurance @accion=5,@estado='"&info(16,2)&"'"
if rs.State=1 then
	rs.Close
end if

Set pcmd = server.CreateObject("ADODB.Command")
pcmd.CommandText = "Insurance..sp_insu_adm_stateInsurance"
pcmd.CommandType = 4 'adCmdStoredProc
pcmd.ActiveConnection = cn
pcmd.Parameters("@accion") = pcmd.CreateParameter("@accion", 131, 1,18, "5") '131=adNumeric, 1=adInput
pcmd.Parameters("@accion").Precision=0
pcmd.Parameters("@accion").NumericScale=18
pcmd.Parameters("@estado") = pcmd.CreateParameter("@estado", 200, 1, 2, info(16,2)) '200=adVarChar, 1=adInput
'pcmd.Parameters.Append(pcmd.CreateParameter("@descripcion", 200, 1, 50, login)) '200=adVarChar, 1=adInput
set rs = pcmd.Execute
'rs.Open misql, cn

if not rs.EOF and not rs.BOF then
	info(16,2)=Rs("descripcion")		  
else
	info(16,2)="Error en tabla estado"
end if
'===================================================================================
'==get info amparo
'===================================================================================
'misql="Insurance..sp_insu_adm_amparo @accion=5,@codamparo='"&info(12,2)&"'"
if rs.State=1 then
	rs.Close
end if

Set pcmd = server.CreateObject("ADODB.Command")
pcmd.CommandText = "Insurance..sp_insu_adm_amparo"
pcmd.CommandType = 4 'adCmdStoredProc
pcmd.ActiveConnection = cn
pcmd.Parameters("@accion") = pcmd.CreateParameter("@accion", 131, 1,18, "5") '131=adNumeric, 1=adInput
pcmd.Parameters("@accion").Precision=0
pcmd.Parameters("@accion").NumericScale=18
pcmd.Parameters("@codamparo") = pcmd.CreateParameter("@codamparo", 200, 1, 3, info(12,2)) '200=adVarChar, 1=adInput
set rs = pcmd.Execute
'rs.Open misql, cn

if not rs.EOF and not rs.BOF then
	info(17,2)=Rs("descripcion")		  
else
	info(17,2)="Error en tabla amparo"
end if
'===================================================================================
'==process info beneficiary
'===================================================================================
'misql="exec Insurance..Beneficiary_Select @producto='"&Request.Form("product")&"',@contrato="&Request.Form("contract")&",@planproducto='"&Request.Form("plan")&"'"
'response.Write misql
'Response.End
if rs.State=1 then
   rs.Close
end if

Set pcmd = server.CreateObject("ADODB.Command")
pcmd.CommandText = "Insurance..Beneficiary_Select"
pcmd.CommandType = 4 'adCmdStoredProc
pcmd.ActiveConnection = cn
pcmd.Parameters("@producto") = pcmd.CreateParameter("@producto", 200, 1, 7, Request.Form("product")) '200=adVarChar, 1=adInput
pcmd.Parameters("@contrato") = pcmd.CreateParameter("@contrato", 131, 1,18, Request.Form("contract")) '131=adNumeric, 1=adInput
pcmd.Parameters("@contrato").Precision=0
pcmd.Parameters("@contrato").NumericScale=18
pcmd.Parameters("@planproducto") = pcmd.CreateParameter("@planproducto", 200, 1, 7, Request.Form("plan")) '200=adVarChar, 1=adInput
'pcmd.Parameters.Append(pcmd.CreateParameter("@fechainicio", 200, 1, 10, begindate)) '200=adVarChar, 1=adInput
set rs = pcmd.Execute
'rs.Open misql, cn

resumen=""
if not rs.EOF and not rs.BOF then
  do while not rs.EOF
    '[0] nombres [1] identificacion [4] tipo de id [2] parentesco [3] porcentaje
	resumen=resumen&rs(0)&";"&rs(1)&";"&rs(4)&";"&rs(2)&";"&rs(3) & "["
	rs.MoveNext  
  loop	
  'remove last [
  resumen=mid(resumen,1,len(resumen)-1)
end if
'===================================================================================
'==flow html
'===================================================================================
OpenHTML
  OpenHead
 	 PlaceLink "REL", "stylesheet", "../../css/style1.css", "text/css"
 	 PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	
	OpenBody "''", "bgcolor='#FFFFFF' text='#000000'"
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
		
		openForm "theform","post",""," "

infoseguro=true
ocultos="N"		
%>
<!--#include file='../pages/displayinsurance.asp'-->
<!--#include file='../pages/displaybenefit.asp'-->

<%

  		  response.Write "<br><br>"
		CloseForm
CloseBody
CloseHTML
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>