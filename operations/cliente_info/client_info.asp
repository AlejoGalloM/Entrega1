<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:				client_info.asp 1300
'Path:					client_info/
'Created By:				A. Orozco 2001/07/30
'Last Modified:				Fabio Calvache Julio 31 2003
'					R. Lagos 2002/07/05
'                                       Add city info an sures millas info
'              				A. Orozco 2001/09/21
'					A. Orozco 2001/10/08
'					Guillermo Aristizabal 2001/10/11
'					A. Orozco 2002/01/08
'Modified by:                           Armando J. Arias Gómez - 2008/05/08 - PlaceTitle/Cambio ID write_sp_log
'Modifications:				File Creation
'Parameters:				Client's ID
'Returns:				Complete client information
'Additional Information:	
'===================================================================================
Option Explicit
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
<%
Authorize 5,2
Dim objConn 'Database Connection
Dim objRst 'Recordset object
Dim strSQL 'Query container
Dim arrClient, arrContracts, arrEmploy, city
Dim strProds, arrProds
Dim I, J, Miles
Dim Contract, Product, Plan, ClientId, DocType
Dim conexion

Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
ClientId = Request.Form("ClientId")
DocType = Request.Form("DocType")
City = Session.Contents("ClientCity")

Dim IsCoreDescription, arrComplement
IsCoreDescription = "No se Conoce"

Set objConn = GetConnPipelineDB
Set objRst = Server.CreateObject("ADODB.Recordset")

' BEGIN: Consulta datos de cliente
strSQL = "spsp_GetClientInfo " & ClientId & ", '" & DocType & "'"
conexion= strSQL

objRst.Open strSQL, objConn
If objRst.BOF And objRst.EOF Then
	arrClient = 0
Else
	arrClient = objRst.GetRows
End If
objRst.Close

write_sp_log objConn, 13180, "spsp_GetClientInfo", Contract, Product, Plan, ClientId, 0, "", "client_info.asp " & _
"Loaded by " & Session("sp_miLogin")
' END: Consulta datos de cliente

' BEGIN: Consulta contratos
strSQL = "sppl_QueContratosTiene " & ClientId & ",'" & Doctype & "'"
conexion= conexion&" - "&strSQL
objRst.Open strSQL, objConn
If objRst.BOF And objRst.EOF Then
	arrContracts = 0
Else
	For I = 0 To objRst.Fields.Count - 1
		strProds = strProds & objRst.Fields(I).Name
		If I < objRst.Fields.Count - 1 Then
			strProds = strProds & ","
		End If
	Next
	arrProds = Split(strProds, ",")
	arrContracts = objRst.GetRows
End If
objRst.Close

write_sp_log objConn, 13180, "sppl_QueContratosTiene", Contract, Product, Plan, ClientId, 0, "", "client_info.asp " & _
"Loaded by " & Session("sp_miLogin")
' END: Consulta contratos

' BEGIN: Consulta empleador
strSQL = "sppl_EmpleadorxCliente " & ClientId & ",'" & Doctype & "'"
conexion= conexion&" - "&strSQL
objRst.Open strSQL, objConn
If objRst.BOF And objRst.EOF Then
	arrEmploy = 0
Else
	arrEmploy = objRst.GetRows
End If
objRst.Close

write_sp_log objConn, 13180, "sppl_EmpleadorxCliente", Contract, Product, Plan, ClientId, 0, "", "client_info.asp " & _
"Loaded by " & Session("sp_miLogin")
' END: Consulta empleador

' BEGIN: Consulta Segmento
strSQL = "spsp_ComplemnetData_GetByClient '" & CStr(DocType) & "', " & CStr(ClientId)
conexion= conexion&" - "&strSQL
objRst.Open strSQL, objConn

If objRst.BOF And objRst.EOF Then
	arrComplement = 0
	IsCoreDescription = "No se Conoce"
Else
	arrComplement = objRst.GetRows()

	If IsArray(arrComplement) Then
		IsCoreDescription = arrComplement(0,0)	
	End If
End If
objRst.Close

write_sp_log objConn, 13180, "Segmento : " & CStr(ClientId) + ":" & DocType, Contract, Product, Plan, ClientId, 0, "", "client_info.asp " & _
"Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"client_info.asp","Segmento : " & CStr(ClientId) + ":" & DocType,Session.contents("name"),conexion ,"N/A","null","Consulta","N/A"
' END: Consulta Segmento

CloseConnPipelineDB


OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
%>
<SCRIPT LANGUAGE=javascript src='../_pipeline_scripts/validation.js'></SCRIPT>
<%
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
                PlaceTitle "Información cliente"
	CloseHead
	OpenBody "", ""
OpenTable "", ""
	OpenTr ""
		OpenTd "thead", ""
			Response.Write "&nbsp;"
		CloseTd
	CloseTr
CloseTable
If IsArray(arrClient) Then
	OpenTable "90%", "'' align=center"
		OpenTr ""
			OpenTd "thead", "colspan=2 align=center"
				Response.Write "<h3>Información del Inversionista</h3>"
			CloseTd
		CloseTr
	CloseTable
	OpenTable "90%", "'' border=1 align=center"
		OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "Nombre"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(0,0)
			CloseTd
		CloseTr
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "No. Identificación"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(1,0)
			CloseTd
		CloseTr
		OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "Tipo de Identificación"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(16,0)
			CloseTd
		CloseTr
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "Dirección"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(2,0)
			CloseTd
		CloseTr
		OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "Ciudad"
			CloseTd
			OpenTd "tbody", ""
				'Response.Write city
				Response.Write arrClient(17,0)
			CloseTd
		CloseTr
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "Teléfono"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(3,0)
			CloseTd
		CloseTr
		OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "Email"
			CloseTd
			OpenTd "tbody", ""
				PlaceAnchor "mailto:" & arrClient(4,0), arrClient(4,0)
			CloseTd
		CloseTr
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "Envío Comunicaciones"
			CloseTd
			OpenTd "tbody", ""
				Select Case UCase(arrClient(5,0))
					Case "T"
						Response.Write "Tradicional"
					Case "E"
						Response.Write "Electrónico"
					Case "B"
						Response.Write "Tradicional y Electrónico"
					Case "N"
						Response.Write "No Enviar"
					Case Else
						Response.Write "N/A"
				End Select
			CloseTd
		CloseTr
		OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "Fecha de Nacimiento"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(8,0)
			CloseTd
		CloseTr
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "Género"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(9,0)
			CloseTd
		CloseTr
        OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "Segmento"
			CloseTd
			OpenTd "tbody", ""
				Response.Write IsCoreDescription
			CloseTd
		CloseTr
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "Metaname"
			CloseTd
			OpenTd "tbody", ""
				Response.Write arrClient(10,0)
			CloseTd
		CloseTr
	 CloseTable
Else
	OpenTable "90%", "'' border=1 align=center"	
		OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "No hay información"
			CloseTd
		CloseTr
	CloseTable
End If
OpenTable "", ""
	OpenTr ""
		OpenTd "thead", ""
			Response.Write "&nbsp;"
		CloseTd
	CloseTr
CloseTable
If IsArray(arrEmploy) Then
	OpenTable "90%", "'' border=1 align=center"
		OpenTr "class=teven"
			OpenTd "thead", "colspan=5"
				Response.Write "Relación de Empleadores"
			CloseTd
		CloseTr
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "Documento"
			CloseTd
			OpenTd "thead", ""
				Response.Write "Empresa"
			CloseTd
			OpenTd "thead", ""
				Response.Write "Dirección"
			CloseTd
			OpenTd "thead", ""
				Response.Write "Ciudad"
			CloseTd
			OpenTd "thead", ""
				Response.Write " Teléfono"
			CloseTd
		CloseTr
		For J = 0 To UBound(arrEmploy, 2)
			If J Mod 2 = 0 Then
				OpenTr "class=todd"
			Else
				OpenTr "class=teven"
			End If
				OpenTd "tbody", ""
					Response.Write arrEmploy(0, J)
				CloseTd
				OpenTd "tbody", ""
					Response.Write arrEmploy(2, J)
				CloseTd
				OpenTd "tbody", ""
					If IsNull(Trim(arrEmploy(3, J))) Or Trim(arrEmploy(3, J)) = "" Then
						Response.Write "&nbsp;"
					Else
						Response.Write Trim(arrEmploy(3, J))
					End If
				CloseTd
				OpenTd "tbody", ""
					If IsNull(Trim(arrEmploy(4, J))) Or Trim(arrEmploy(4, J)) = "" Then
						Response.Write "&nbsp;"
					Else
						Response.Write arrEmploy(4, J)
					End If
				CloseTd
				OpenTd "tbody", ""
					If IsNull(Trim(arrEmploy(6, J))) Or Trim(arrEmploy(6, J)) = "" Then
						Response.Write "&nbsp;"
					Else
						Response.Write arrEmploy(6, J)
					End If
				CloseTd
			CloseTr
		Next
	CloseTable
Else
	OpenTable "90%", "'' border=1 align=center"
		OpenTr "class=teven"
			OpenTd "thead", ""
				Response.Write "No tiene empleadores registrados"
			CloseTd
		CloseTr
	CloseTable
End If
	OpenTable "", ""
		OpenTr ""
			OpenTd "thead", ""
				Response.Write "&nbsp;"
			CloseTd
		CloseTr
	CloseTable
If IsArray(arrContracts) Then
	OpenTable "90%", "'' border=1 align=center"
		OpenTr "class=teven"
			OpenTd "thead", "colspan=" & UBound(arrContracts) + 1
				Response.Write "Contratos del inversionsta en todos los productos"
			CloseTd
		CloseTr
		OpenTr "class=teven"
			For I = 0 To UBound(arrProds) '- 1
				OpenTd "thead", "align=center"
					Response.Write arrProds(I)
				CloseTd
			Next
		CloseTr
		OpenTr "class=todd"
			For I = 0 To UBound(arrContracts)' - 1
				OpenTd "tbody", "align=center"
					Response.Write arrContracts(I, 0)
				CloseTd
			Next
		CloseTr
	CloseTable
Else
	OpenTable "", ""	
		OpenTr "class=todd"
			OpenTd "thead", ""
				Response.Write "No tiene contratos"
			CloseTd
		CloseTr
	CloseTable
End If
CloseBody
CloseHTML

If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
