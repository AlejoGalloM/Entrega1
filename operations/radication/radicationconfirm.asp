<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		radication.asp 4101
'Path:			/operations/radicationConfirm.asp
'Created By:		Guillermo Pinerez 2001/08/22
'Last Modified:		J M Moreno 2003/11/20 modify buttons radication removed enviar radicacion.
'			A. Orozco 2001/10/08
'			Guillermo Aristizabal 2001/10/11
'                       Armando Arias - 2008/May/09 - Cumplimiento circular SF052 - PlaceTitle - write_sp_log 13304
'Parameters:						Contract
'							Unit
'							ClientId
'							Product
'							AgentId
'Returns:
'Additional Information:	
'===================================================================================
Option Explicit
On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../_pipeline_scripts/tags.asp"-->
<!--#include file="../_pipeline_scripts/pipeline_scripts.asp"-->
<!--#include file="../../operations/transactions/reglamentoscriptsPipeline.asp"-->
<%
Authorize 1,6
Response.Write "<link rel='stylesheet' href='../../css/style.css' type='text/css'>" & vbCrLf

'== declares ===
Dim Product
Dim Contract
Dim Plan
dim nroDocum
dim DocType
dim Name
dim ProcessName
dim arrRadication
dim arrCliente
dim arrProcess
Dim objConn
Dim Proceso
Dim processInfo, component_id, value

dim Reference '<I&T - DMPC>

function getProcessName(idProcess)
dim arrProcess
dim J
arrProcess = Application.Contents("Process")

	getProcessName = "error..."
	
	If IsArray(arrProcess) Then
		For J = 0 To UBound(arrProcess, 2)'rows
			if cint(arrProcess(0,J)) = cint(idProcess) then
				getProcessName = arrProcess(1,J)
				exit function
			end if
		next		
	else
		getProcessName =  "Error: ID de proceso no valido"
	end if
	
end function

function getTipoFranquicia(contrato)
dim adoConn
dim strSQL
dim objRst
dim arrProcess
on error resume next
	set adoConn = GetConnpipelineDB()
	Set objRst = Server.CreateObject("ADODB.Recordset")
	strSQL = "exec sprd_getSocietyType " & cstr(contrato) & ",'" & request.form("Product") & "'"
	objRst.Open strSQL, adoConn
	
        write_sp_log adoConn, 13304, "sprd_getSocietyType", contrato, request.form("Product") , "", 0, 0, "", "radicationconfirm.asp " & _
        " Loaded by " & Session("sp_miLogin")

		component_id = "radicationconfirm.asp"
		processInfo =  "radicationconfirm.asp " & " Loaded by " & Session("sp_miLogin")

	If objRst.BOF And objRst.EOF Then
		getTipoFranquicia = "Indefinida"
	Else
		arrProcess = objRst.GetRows
		if IsArray(arrProcess) then
			getTipoFranquicia = arrProcess
		else
			getTipoFranquicia = "Error en campo area"
		end if
	End If
	objRst.Close
	exit function	
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
end function

function getProcessidArea(idProcess)
dim arrProcess
dim J
arrProcess = Application.Contents("Process")

	getProcessidArea = "error..."
	
	If IsArray(arrProcess) Then
		For J = 0 To UBound(arrProcess, 2)'rows
			if cint(arrProcess(0,J)) = cint(idProcess) then
				getProcessidArea = arrProcess(4,J)
				exit function
			end if
		next		
	else
		getProcessidArea =  "Error: ID de proceso no valido"
	end if
end function


function getSurrenderName(surrenderType)
 if surrenderType = 1 then
	getSurrenderName = "PS Partial surrender"
 else
	if surrenderType = 2 then
		getSurrenderName = "FS Full surrender"
	else
		if surrenderType = 3 then
			getSurrenderName = "T Traslado"
		else
			getSurrenderName = " " '& surrenderType
		end if
	end if
 end if
end function

function getProcessArea(idProcess, contrato)
dim arrProcess
dim J
dim arrAreas
arrProcess = Application.Contents("Process")

	getProcessArea = "error..."
	
	If IsArray(arrProcess) Then
		For J = 0 To UBound(arrProcess, 2)'rows
			if cint(arrProcess(0,J)) = cint(idProcess) then
				if trim(LCase( arrProcess(3,J))) = "franquicia" then
					if contrato <> "" then
						getProcessArea = getTipoFranquicia(contrato)
					end if
				else
					if contrato <> "" then
						arrAreas = getTipoFranquicia(contrato)
					end if
					if IsArray(arrAreas) then
						arrAreas(0,0) = trim(LCase( arrProcess(3,J)))
						getProcessArea = arrAreas					
					else
						dim arrAreasx(1,1)
						arrAreasx(0,0) = trim(LCase( arrProcess(3,J)))
						arrAreasx(1,0) = "No tiene unidad"
						getProcessArea = arrAreasx
					end if
				end if
				exit function
			end if
		next		
	else
		getProcessArea =  "Error: ID de proceso no valido"
	end if
end function


Function ReceptionWayName(idRecWay)
dim Combo
dim arrProcess
dim J 


if len(idRecWay) = 0 then
	ReceptionWayName = "No lleno la forma de recepcion..."
	exit function
end if
if len( Application.Contents("receptionWay") ) = 0 then
	arrProcess = getReceptionWay()
else
	arrProcess = Application.Contents("receptionWayArr")
end if
		For J = 0 To UBound(arrProcess, 2)'rows
			if cint( arrProcess(0,J)) = cint(idRecWay)  then
				ReceptionWayName = arrProcess(1,J)
			end if
		next	

end function 


'Response.Write "El proceso es " & request.form("Process")
proceso = request.form("Process")
Set objConn = GetConnPipelineDB

'write_sp_log objConn, 4101, "", 0, "" , "", 0, 0, "", "radicationconfirm.asp " & _
'" Loaded by " & Session("sp_miLogin")
write_sp_log objConn, 13304, "", 0, "" , "", 0, 0, "", "radicationconfirm.asp " & _
" Loaded by " & Session("sp_miLogin")

CloseConnPipelineDB

Contract = request.form("contract")
'==========================================================================
''<I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia �nica a partir del Contrato y el producto
'==========================================================================
 Reference = GetReferenciaUnica( request.form("Product"), request.form("contract"))
'==========================================================================

OpenTable "70%", "'' align=center border=0"
		OpenTr ""
			OpenTd "thead", "align=center "
				Response.Write "Confirmaci�n de datos de la solicitud"
			CloseTd
		CloseTr
'== START customer Name ==
		OpenTr ""
			OpenTd "todd", "align=left"
				Response.Write "Nombre del cliente "
			CloseTd
			OpenTd "todd", "align=left "			
				Response.Write request.form("ClientName")
				value = "Nombre del cliente: "&SubstitutePlaceholders(request.form("ClientName"))
			
			CloseTd
		CloseTr
'== END customer name ==
'== START customer ID ==
		OpenTr ""
			OpenTd "teven", "align=left"
				Response.Write "Identificaci�n "
			CloseTd
			OpenTd "teven", "align=left "			
				Response.Write request.form("ClientId")
				Response.Write " "
				Response.Write request.form("DocType")
				value = value&", Identificación: "&SubstitutePlaceholders(request.form("ClientId"))
			
			CloseTd
		CloseTr
'== END customer ID ==
'== START contract ID ==
		OpenTr ""
			OpenTd "todd", "align=left"
				Response.Write "Contrato "
			CloseTd
			OpenTd "todd", "align=left "			
				'Response.Write request.form("contract") '<I&T - DMPC - Modificado por Proceso de Referencia Unica - Recaudos>
				Response.Write Reference
				value = value&", Contrato: "&Reference
			CloseTd
		CloseTr
'== END customer ID ==
'== START Product ID ==
		OpenTr ""
			OpenTd "teven", "align=left"
				Response.Write "Producto "
			CloseTd
			OpenTd "teven", "align=left "			
				Response.Write request.form("Product")
				value = value&", Producto: "&SubstitutePlaceholders(request.form("Product"))
			
			CloseTd
		CloseTr
'== END Product ID ==
'== START Process ID ==
		OpenTr ""
			OpenTd "todd", "align=left"
				Response.Write "Proceso "
			CloseTd
			OpenTd "todd", "align=left "			
				Response.Write getProcessName(Proceso)
				value = value&", Proceso: "& getProcessName(Proceso)
			CloseTd
		CloseTr
'== END Process ID ==
'== START reception way ==
		OpenTr ""
			OpenTd "teven", "align=left"
				Response.Write "Forma recepci�n "
			CloseTd
			OpenTd "teven", "align=left "			
				Response.Write ReceptionWayName(request.form("ReceptionWay"))
				value = value&", Forma de recepción: "&ReceptionWayName(SubstitutePlaceholders(request.form("ReceptionWay")))
			
			CloseTd
		CloseTr
'== END  reception way ==
'== START Process ID ==
		OpenTr ""
			OpenTd "todd", "align=left"
				Response.Write "Valor retiro"
			CloseTd
			OpenTd "todd", "align=left "
				if IsNumeric(request.form("surrendervalue"))  then
					Response.Write FormatCurrency(request.form("surrendervalue"),2)
					value = value&", Valor Retiro: "&FormatCurrency(SubstitutePlaceholders(request.form("surrendervalue")),2)
				
				else
					Response.Write (request.form("surrendervalue"))
					value = value&", Valor Retiro: "&(SubstitutePlaceholders(request.form("surrendervalue")))
				
				end if	
			CloseTd
		CloseTr
'== END Process ID ==
'== START surrender type ==
		OpenTr ""
			OpenTd "teven", "align=left"
				Response.Write "Tipo retiro"
			CloseTd
			OpenTd "teven", "align=left "
				Response.Write getSurrenderName( request.form("surrrendertype"))
				value = value&", Tipo de Retiro: "&getSurrenderName( SubstitutePlaceholders(request.form("surrrendertype")))
			
			CloseTd
		CloseTr
'== END surrender type ==
'== START comments ==
		OpenTr ""
			OpenTd "todd", "align=left"
				Response.Write "Comentarios"
			CloseTd
			OpenTd "todd", "align=left "
				Response.Write request.form("Comments")
				value = value&", Comentarios: "&SubstitutePlaceholders(request.form("Comments"))
			
			CloseTd
		CloseTr
'== END surrender type ==
'== START comments ==
		OpenTr ""
			OpenTd "teven", "align=left"
				Response.Write "<label class=wflowarea>Area, Unidad</label>"
			CloseTd
			OpenTd "teven", "align=left "
				dim arrAreaUnidad
				dim area,unidad
				Response.Write "<label class=wflowarea>"
				arrAreaUnidad = getProcessArea(Proceso, Contract ) 
				if IsArray(arrAreaUnidad) then
					area = arrAreaUnidad(0,0)
					unidad = arrAreaUnidad(1,0)
				else
					area = "no definida"
					unidad = "Unidad no definida"
				end if
					Response.Write cstr(area)
					Response.Write ", "
					Response.Write cstr(unidad)
				Response.Write " </label>"
			CloseTd

		write_dataLog Response.Status,component_id,processInfo,Session.contents("name"), "sprd_getSocietyType" ,"",value,"Operación-Adición","N/A"

		CloseTr
'== END surrender type ==
CloseTable

OpenTable "40%", "'' align=center border=0"
		OpenTr ""
			OpenTd "todd", "align=center"				
				OpenForm "PBR", "Post", "Radication.asp", ""
				PlaceInput "idArea", "hidden", Area, ""
				PlaceInput "Area", "hidden", Area, ""
				PlaceInput "Unidad", "hidden", Unidad, ""

				dim key
				For Each key in Request.Form 
					PlaceInput (Key), "hidden", Request.Form(Key), ""
				Next 
				PlaceInput "Modify", "Submit", "modificar Radicaci�n", "class=sbttn"
				CloseForm				
			CloseTd
			OpenTd "todd", "align=center"				
				OpenForm "PB", "Post", "SaveRadication.asp?repeat=no", ""
				PlaceInput "Area", "hidden", Area, ""
				PlaceInput "Unidad", "hidden", Unidad, ""
				PlaceInput "idArea", "hidden", cstr(getProcessidArea(Proceso)), ""				
				
				For Each key in Request.Form '
					PlaceInput (Key), "hidden", Request.Form(Key), ""
				Next 
				PlaceInput "SaveR", "Submit", "Enviar Radicaci�n", "class=sbttn"
				CloseForm
			CloseTd
			OpenTd "todd", "align=center "
				OpenForm "PBS", "Post", "SaveRadication.asp?repeat=yes", ""
				PlaceInput "Area", "hidden", Area, ""
				PlaceInput "Unidad", "hidden", Unidad, ""
				PlaceInput "idArea", "hidden", cstr(getProcessidArea(request.form("Process"))), ""				
				
				For Each key in Request.Form '
					PlaceInput (Key), "hidden", Request.Form(Key), ""
				Next 
				PlaceInput "SaveS", "Submit", "Enviar y Radicar al mismo Cliente", "class=sbttn"
				CloseForm
			CloseTd
		CloseTr
CloseTable
%>