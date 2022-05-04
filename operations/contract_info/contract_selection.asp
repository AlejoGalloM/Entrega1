<%@ Language=VBScript %>
<%


'===================================================================================
'File Name:		contract_selection.asp 1400
'Path:				contract_info/
'Created By:		A. Orozco 2001/07/25
'Last Modified:		J. moreno 2003/10/30 Add menu productos prototipo fco, fibac igold tlife fonvida fmagno
'				J carreño 2003/09/15 Add Menu Alternativo
'				Fabio Calvache Agosto 5/2003 Add DocType Marathon
'				R. Lagos 2002/11/15 Remove Pas reference in  case 2
'				R. Lagos 2002/02/04  Add Pas
'				A. Orozco 2001/09/10
'				Guillermo Aristizabal  2001/09/18 auth & log
'				A. Orozco 2001/10/08
'				Guillermo Aristizabal 2001/10/11
'				A. Orozco 2001/10/30
'				A. Orozco 2001/11/26
'				A. Orozco 2001/12/07
'				R. Lagos 2002/02/01  Add sistem PAS
'				javier VArgas 2005/30/06 add web services solution
'				Camilo Gutierrez I&T 11-02-2011 Add Risk Profile( Perfil del Inversionista ) & Real Contract Profile ( Portafolio de Inversion ) Columns
'Parameters:		User must be logged on
'						Session("nrodocum")
'Returns:			List of active contracts for the client
'					   Oscar Diaz 2012/11/06
'					   Se utiliza una función para comprobar los planes corporativos
'					   y evitar el uso de variables quemadas en el asp
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
<!--#include file="../_pipeline_scripts/mfundCorporativoScripts.asp"-->
<!--#include file="../../operations/transactions/reglamentoscriptsPipeline.asp"-->
<%
Authorize 6,2
Dim I, J 'Contract Counter
Dim cn, rs 'Connection and Recordset
Dim Sql	'String for the Stored Procedure
Dim strSql 'String for the Stored Procedure
Dim arrAuxiliar 'Recordset Data for active contracts
Dim Product, Contract, Plan, Name, ClientId, DocType, Phone 'Store Product and Contract ID
Dim Total 'Contracts Total
Dim bc, fechaCierre, Socs, arrSocs, Flag, K, Pas
dim Reference  ' <I&T - DMPC>
dim conexion

'Jose Alejandro Figueroa - Proyecto Advice Tools - 2011-03-16
dim hasContractInfo
hasContractInfo = false


Contract = Request.Form("Contract")
Product = Request.Form("Product")
Plan = Request.Form("Plan")
Name = Request.Form("Name")
ClientId = Request.Form("ClientId")
DocType = Request.Form("DocType")
'Response.Write "Manito : "
'Response.Write DocType
Phone = Request.Form("Phone")





'Get Client's Active Contracts
'Sql = "Exec spsp_GetClientContracts " & CStr(Session.Contents("sp_nrodocum")  )
Sql = "Exec spsp_GetClientContracts " & CStr(ClientId) & ", '" & DocType & "'"
conexion="spsp_GetClientContracts " & CStr(ClientId) & ", '" & DocType & "'"
Set cn = GetConnpipelineDB
Set rs=server.CreateObject("ADODB.Recordset")
rs.Open sql,cn
If rs.BOF And rs.EOF Then
	arrAuxiliar = 0
Else
	arrAuxiliar = rs.GetRows() 
End If
rs.Close

write_sp_log cn, 1400, "spsp_GetClientContracts", Contract, Product, Plan, ClientId, 0, "", "contract_selection.asp " & _
"- " & Session("sp_miLogin")



'Get whole saler's societies
If Session("sp_AccessLevel") = "1" Then
	Sql = "spsp_GetWSSocieties '0" & Session("sp_IdAgte") & "'"
	conexion=conexion&" - "&Sql
	rs.Open Sql, cn
	If rs.BOF And rs.EOF Then
		Socs = 0
	Else
		arrSocs = rs.GetRows()
	End If
	rs.Close
	write_sp_log cn, 1400, "spsp_GetWSSocieties", Contract, Product, Plan, ClientId, 0, "", "contract_selection.asp " & _
	"- " & Session("sp_miLogin")
End If

'Set rs = Nothing
'CloseConnpipelineDB
'Set cn = Nothing

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
set bc = Server.CreateObject("MSWC.BrowserType")

%>
<SCRIPT LANGUAGE="javascript">
<!--
function Reload_RightFrame(Params) {
	window.parent.parent.frames(3).location = "../menus/menu_right.asp?" + Params;
}
function submitter(cont, prod, plan, Pas, url, clientid, name, phone, docType) {
	document.selection.elements['Contract'].value=cont;
	document.selection['Product'].value=prod;
	document.selection['Plan'].value=plan;
	document.selection['Pas'].value=Pas;
	document.selection['ClientId'].value=clientid;
	document.selection['DocType'].value=docType;
	document.selection['Name'].value=name;
	document.selection.action="ChangeContract.asp?Url="+url;
	document.menu.elements['Contract'].value=cont;
	document.menu['Product'].value=prod;
	document.menu['Plan'].value=plan;
	document.menu['Pas'].value=plan;
	document.menu['ClientId'].value=clientid;
	document.menu['DocType'].value=docType;
	document.menu['Name'].value=name;
	document.menu['Phone'].value=phone;
	document.menu.submit();
	document.selection.submit();
	return false
}
//-->
</SCRIPT>
<%

%>
<SCRIPT LANGUAGE="javascript">
<!--
//function Reload_RightFrame(Params) {
//	top.frames.frames[3].location = ""../menus/menu_right.asp?"" + Params;
//}
//-->
</SCRIPT>
<%
'End If
'------------------------------------------------------
'Finaliza Javascript para recargar el menu de la derecha
'------------------------------------------------------
	CloseHead
	OpenBody "", ""




If IsArray(arrAuxiliar) Then
'Display Active Contracts Information (Table Header)
Response.Write "<p>&nbsp;</p><p align=center>"
Response.Write "</p><p>&nbsp;</p>" & vbCrLf
OpenTable "70%", "'' border=1 align=center"
	OpenTr "class=teven"
		OpenTd "thead", "align=center"
			Response.Write "Producto/Servicio"
		CloseTd
		OpenTd "thead", "align=center"
			'Response.Write "No. de afiliación" '<I&T - DMPC - Modificado por proceso de Referencia Unica - Recaudos>
				Response.Write "No. de Contrato"
		CloseTd
		OpenTd "thead", "align=center"
			Response.Write "Estado"
		CloseTd
		OpenTd "thead", "align=center"
			Response.Write "Perfil de inversión del contrato"
		CloseTd
		OpenTd "thead", "align=center"
			Response.Write "Portafolio del inversionista"
		CloseTd
		OpenTd "thead", "align=center"
			Response.Write "Saldo"
		CloseTd
	CloseTr




	'Build Table to display (Table Body)
	For J = 0 To UBound(arrAuxiliar, 2) 'Rows
		If (J Mod 2) = 0 Then
			OpenTr "class=todd"
		Else
			OpenTr "class=teven"
		End If
		
		'Alejandro Figueroa - Proyecto Advice Tools - 2011-03-16
		'La funcionalidad aplica para todos los contratos Multifund Individual y para los Corporativo tipo CAHC (los demás no)
		Product = arrAuxiliar(0,J)
		Plan = arrAuxiliar(1,J)

		hasContractInfo = false
		if trim(Product) = "MFUND" and  VerificarPlanProducto(Application("CorpPlanesAll"),trim(Plan)) <>  true  Then
			hasContractInfo = true
		End if
		
		For I = 0 To UBound(arrAuxiliar)'Cols
			
			Select Case I
				Case 0 'Product Description
					Product = arrAuxiliar(I,J)
					OpenTd "tbody", "align=center"
						Response.Write Product
				Case 1 'Plan (don't display)
					Plan = arrAuxiliar(I,J)
				Case 2 'Contract and Current Balance
					Contract = arrAuxiliar(I,J)
					
					OpenTd "tbody", "align=center"
						Dim ContractName
						Dim page
						
						pas="N"
						ContractName = ""
						
						Select Case ucase(RTrim(Product))
							Case "MFUND","TLIFE","SIPEN"
								page = "contract_info_mfund.asp"
								
								'======================================================
								'Camilo Gutierrez I&T 10-02-2011 Agregar Nombre de contrato
								'======================================================
								If ucase(RTrim(Product)) = "MFUND" Then
									strSql = "sppl_GetDetallesContrato " & Contract & ", '" & Product & "', '" & Plan & "'"
									conexion=conexion&" - "&strSql
									Dim rstContractInfo
									Dim arrContractInfo
									Set rstContractInfo = Server.CreateObject("ADODB.Recordset")
									rstContractInfo.Open strSql, cn

									If rstContractInfo.BOF And rstContractInfo.EOF Then
										arrContractInfo = 0
									Else
										arrContractInfo = rstContractInfo.GetRows()
										
										If hasContractInfo and not isnull(arrContractInfo(103,0)) and arrContractInfo(103,0) <> "" Then
											ContractName = arrContractInfo(103,0) & " <img title='Su cliente podrá definir esta información a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' /><br/>"
										End If
										
									End If
									rstContractInfo.Close
								End If
								'======================================================
								'End Camilo Gutierrez I&T 10-02-2011 Agregar Nombre de contrato
								'======================================================
								
								If ucase(RTrim(Product))="SIPEN" Then
								         Pas="S"
								Else         
									'page = "contract_info_mfund.asp"
									Sql = "exec sppl_getpas '" & product & "'," & contract
									conexion=conexion&" - sppl_getpas " & product & "'," & contract
						        		'Response.Write Sql
						        		if rs.State=1 then
						           			rs.Close
									end if
									rs.Open sql, cn
									if not rs.EOF and not rs.BOF then
								   		if rs(0)>0 then  'have a PAS
								      			Pas="S"
								   		else
								      			Pas="N"   
								   		end if
									end if
									if rs.State=1 then
								  		rs.Close
									end if    
								end if
							Case "MTCOR", "MTIND"
								page = "contract_info_mtcor.asp"
								Sql = "exec sppl_getpas '" & product & "'," & contract
						        conexion=conexion&" - sppl_getpas " & product & "'," & contract
								'Response.Write Sql
						        if rs.State=1 then
						           rs.Close
								end if
								rs.Open sql, cn
								if not rs.EOF and not rs.BOF then
								   if rs(0)>0 then  'have a PAS
								      Pas="S"
								   else
								      Pas="N"   
								   end if
								end if
								if rs.State=1 then
								  rs.Close
								end if    
								
							Case "FPOB"
								page = "contract_info_fpob.asp"
							'======2003/09/15 J Carreño	
							Case "FPAL"
								page = "contract_info_fpal.asp"
							'=====Fin Modificación	
							'======2003/10/30 J moreno	
							Case "FCO", "TLIFE", "IGOLD", "FONVIDA", "FIBAC", "FMAGNO", "SKINST", "OMBRAV", "OMINMA", "OMACCI", "OMLIQ", "ICGREN","ICINME","ICINMR","ICLIME","ICOPOR","ICRFGO","ICRFLO","ICRVGO","ICT108","ICT187","OMACCP","ICCATI","ICCATII","ICCATVI"
								page = "contract_info_proto.asp"
								''<I&T - DMPC 2009/03/11 - Modificado por proceso de Referencia Unica - Recaudos>
								'if ucase(RTrim(Product))="FCO" then
									
							'=====Fin Modificación		
							Case Else
								page = "contract_info_other.asp"
						End Select
						'==========================================================================
						''<I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica - Recaudos>
						' Obtiene la referencia única a partir del Contrato y el producto
						'==========================================================================
							dim ContratoFCO

							If Trim(Product) = "FCO" Then
								ContratoFCO = Trim(Plan)+ Trim(CStr(Contract))
								Reference = GetReferenciaUnica( Product, ContratoFCO)
							else
								Reference = GetReferenciaUnica( Product, Contract)
							end if
							
						'==========================================================================
			
						Select Case Session("sp_AccessLevel")
							Case 0 'Skandia Worker
								'<I&T - DMPC - Modificado por Proceso de REferencia Unica - Recaudos>
								'PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
								'Trim(Product) & """,""" & Trim(Plan)  & """,""" & Trim(Pas)  & """,""" & Trim(Page) & """, """ & Trim(ClientId) & _
								'""", """ & Trim(Name) & """, """ & Trim(Phone)& """, """ & Trim(DocType) & """" & ")' """, Contract 
								if hasContractInfo Then
									Response.Write ContractName
								End if
								
								PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
								Trim(Product) & """,""" & Trim(Plan)  & """,""" & Trim(Pas)  & """,""" & Trim(Page) & """, """ & Trim(ClientId) & _
								""", """ & Trim(Name) & """, """ & Trim(Phone)& """, """ & Trim(DocType) & """" & ")' """, Reference 

							Case 1 ' Whole Saler
								Flag = False
								For K = 0 To UBound(arrSocs,2)
									If Not(IsNull(arrAuxiliar(4,J))) Then
										If CStr(arrSocs(0,K)) = CStr(arrAuxiliar(4,J)) Then
											Flag = True
											Exit For
										Else
											Flag = False
										End If
									End If
								Next
								If Flag Then
								'<I&T - DMPC - Modificado por Proceso de REferencia Unica - Recaudos>
								'	PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
								'	Trim(Product) & """,""" & Trim(Plan)  & """,""" & Trim(Pas)  & ""","""& Trim(Page) & """, """ & Trim(ClientId) & _
								'	""", """ & Trim(Name) & """, """ & Trim(Phone) & """, """ & Trim(DocType) & """" & ")' """, Contract 
									
									PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
									Trim(Product) & """,""" & Trim(Plan)  & """,""" & Trim(Pas)  & ""","""& Trim(Page) & """, """ & Trim(ClientId) & _
									""", """ & Trim(Name) & """, """ & Trim(Phone) & """, """ & Trim(DocType) & """" & ")' """, ContractName & Reference 
								Else
									'Response.Write Contract ' <I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica>
									Response.Write (Reference)
								End If
							Case 2 'Partner FP, Fran. Worker
								If Not IsNull(arrAuxiliar(4,J)) Then
									If CStr(Session("sp_idSoc")) = CStr(arrAuxiliar(4,J)) Then
									'<I&T - DMPC - Modificado por Proceso de REferencia Unica - Recaudos>
									'	PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
									'	Trim(Product) & """,""" & Trim(Plan) & """,""" & Trim(Pas)  & ""","""& Trim(Page)  & """, """ & Trim(ClientId) & _
									'	""", """ & Trim(Name) & """, """ & Trim(Phone) & """, """ & Trim(DocType) & """" & ")' """, Contract ' & """,""" & Trim(Pas) & """, """
										
										PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
										Trim(Product) & """,""" & Trim(Plan) & """,""" & Trim(Pas)  & ""","""& Trim(Page)  & """, """ & Trim(ClientId) & _
										""", """ & Trim(Name) & """, """ & Trim(Phone) & """, """ & Trim(DocType) & """" & ")' """, ContractName & Reference ' & """,""" & Trim(Pas) & """, """
									Else
										'Response.Write Contract ' <I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica>
										Response.Write (Reference)
									End If
								Else
									'Response.Write Contract ' <I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica>
									Response.Write (Reference)
								End If
							Case 3 'FP
								If CStr(Session("sp_IdAgte")) = CStr(arrAuxiliar(5,J)) Then
								'<I&T - DMPC - Modificado por Proceso de REferencia Unica - Recaudos>
								'	PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
								'	Trim(Product) & """,""" & Trim(Plan)  & """, """ & Trim(Pas)  & ""","""& Trim(Page) & """, """ & Trim(ClientId) & _
								'	""", """ & Trim(Name) & """, """ & Trim(Phone) & """, """ & Trim(DocType) & """" & ")' """, Contract
									
									PlaceAnchor """ onClick='return submitter(""" & Trim(Contract) & """,""" & _
									Trim(Product) & """,""" & Trim(Plan)  & """, """ & Trim(Pas)  & ""","""& Trim(Page) & """, """ & Trim(ClientId) & _
									""", """ & Trim(Name) & """, """ & Trim(Phone) & """, """ & Trim(DocType) & """" & ")' """, ContractName & Reference
								Else
									'Response.Write Contract ' <I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica>
									Response.Write (ContractName & Reference)
								End If
							Case 4' Anyone else
						End Select
					CloseTd
					OpenTd "tbody", "align=center"
						If IsNull(arrAuxiliar(6,J)) Then
							Response.Write "&nbsp;" & vbCrLf
						Else
							Response.Write arrAuxiliar(6,J)
						End If
					CloseTd


					'======================================================
					'Camilo Gutierrez I&T 10-02-2011 Agregar Columna Risk Profile
					'======================================================
					OpenTd "tbody", "align=center"
						if hasContractInfo Then
							strSql = "Relacionamiento..Contract_Profile_GetProfileByContractByCountry '" & Trim(DocType) & Trim(ClientId) & "', 'CO', " & Trim(Contract) & ", '" & Trim(Product) & "'"
							conexion=conexion&" - "&strSql
							Dim rstRiskProfile
							Dim arrRiskProfile
							Dim strRiskProfile
							Set rstRiskProfile = Server.CreateObject("ADODB.Recordset")
							rstRiskProfile.Open strSql, cn

							If rstRiskProfile.BOF And rstRiskProfile.EOF Then
								arrRiskProfile = 0
								strRiskProfile = "No definido <img title='Este perfil es definido por su cliente en la Encuesta de Perfil de Inversión del formato de vinculación Multifund o a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' tag='" & Contract & "' />"
							Else
								arrRiskProfile = rstRiskProfile.GetRows()
								If isnull(arrRiskProfile(2,0)) or arrRiskProfile(2,0) = "" Then
									strRiskProfile = "No definido <img title='Este perfil es definido por su cliente en la Encuesta de Perfil de Inversión del formato de vinculación Multifund o a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' tag='" & Contract & "' />"
								Else
									strRiskProfile = arrRiskProfile(2,0) & "  <img title='Este perfil es definido por su cliente en la Encuesta de Perfil de Inversión del formato de vinculación Multifund o a través del Portal de Clientes' src='../../images/operations/InfoIcon.png' tag='" & Contract & "' />"
								End If
							End If
							Response.Write  strRiskProfile
							rstRiskProfile.Close
						Else
							Response.Write "N/A"
						End If
					CloseTd
					'======================================================
					'End Camilo Gutierrez I&T 10-02-2011 Agregar Columna Risk Profile
					'======================================================
					
					'======================================================
					'Camilo Gutierrez I&T 10-02-2011 Agregar Columna Real Profile
					'======================================================
					OpenTd "tbody", "align=center"
						if hasContractInfo Then
							strSql = "Relacionamiento..Contract_Profile_GetRealRiskProfileByContractByCountry '" & Trim(DocType) & Trim(ClientId) & "', 'CO', " & Trim(Contract) & ", '" & Trim(Product) & "'"
							conexion=conexion&" - "&strSql
							Dim rstRealProfile
							Dim arrRealProfile
							Dim strRealProfile
							Set rstRealProfile = Server.CreateObject("ADODB.Recordset")
							rstRealProfile.Open strSql, cn

							If rstRealProfile.BOF And rstRealProfile.EOF Then
								arrRealProfile = 0
								strRealProfile = "No definido <img title='Este perfil es calculado a partir de la distribución real de la inversiones de su cliente en los diferentes portafolios' src='../../images/operations/InfoIcon.png' tag='" & Contract & "' />"
							Else
								arrRealProfile = rstRealProfile.GetRows()
								If isnull(arrRealProfile(2,0)) or arrRealProfile(2,0) = "" Then
									strRealProfile = "No definido <img title='Este perfil es calculado a partir de la distribución real de la inversiones de su cliente en los diferentes portafolios' src='../../images/operations/InfoIcon.png' tag='" & Contract & "' />"
								Else
									strRealProfile = arrRealProfile(2,0) & " <img title='Este perfil es calculado a partir de la distribución real de la inversiones de su cliente en los diferentes portafolios' src='../../images/operations/InfoIcon.png' tag='" & Contract & "' />"
								End If
							End If
							Response.Write  strRealProfile
							rstRealProfile.Close
						Else
							Response.Write "N/A"
						End If
					CloseTd


					'======================================================
					'End Camilo Gutierrez I&T 10-02-2011 Agregar Columna Real Profile
					'======================================================
					OpenTd "tbody", "align=right"
						Select Case Session("sp_AccessLevel")
							Case 0 'Skandia Worker
								If IsNull(arrAuxiliar(3,J)) Then
									Response.Write FormatCurrency(0, 2)
								Else
									Response.Write FormatCurrency(arrAuxiliar(3,J), 2)
									Total = Total + arrAuxiliar(3,J)
								End If
							Case 1 ' Whole Saler
								Flag = False
								For K = 0 To UBound(arrSocs,2)
									If Not(IsNull(arrAuxiliar(4,J))) Then
										If CStr(arrSocs(0,K)) = CStr(arrAuxiliar(4,J)) Then
											Flag = True
											Exit For
										Else
											Flag = False
										End If
									End If
								Next
								If Flag Then
									If IsNull(arrAuxiliar(3,J)) Then
										Response.Write FormatCurrency(0, 2)
									Else
										Response.Write FormatCurrency(arrAuxiliar(3,J), 2)
										Total = Total + arrAuxiliar(3,J)
									End If
								Else
									Response.Write "N/A"
								End If
							Case 2 'Partner FP, Fran. Worker
								If Not IsNull(arrAuxiliar(4,J)) Then
									If CStr(Session("sp_idSoc")) = CStr(arrAuxiliar(4,J)) Then
										If IsNull(arrAuxiliar(3,J)) Then
											Response.Write FormatCurrency(0, 2)
										Else
											Response.Write FormatCurrency(arrAuxiliar(3,J), 2)
											Total = Total + arrAuxiliar(3,J)
										End If
									Else
										Response.Write "N/A"
									End If
								Else
									Response.Write "N/A"
								End If
							Case 3 'FP
								If CStr(Session("sp_IdAgte")) = CStr(arrAuxiliar(5,J)) Then
									If IsNull(arrAuxiliar(3,J)) Then
										Response.Write FormatCurrency(0, 2)
									Else
										Response.Write FormatCurrency(arrAuxiliar(3,J), 2)
										Total = Total + arrAuxiliar(3,J)
									End If
								Else
									Response.Write "N/A"
								End If
							Case 4' Anyone else
									Response.Write "N/A"
						End Select
						'If IsNull(arrAuxiliar(3,J)) Then
						'	Response.Write FormatCurrency(0, 2)
						'Else
						'	Response.Write FormatCurrency(arrAuxiliar(3,J), 2)
						'End If
					CloseTd
						'If IsNull(arrAuxiliar(3,J)) or arrAuxiliar(3,J) = "" Then
						'Else
						'	Total = Total + arrAuxiliar(3,J)
						'End If
			End Select
		Next
		CloseTr
	Next
'-------------------------------------------------	
Set rs = Nothing
CloseConnpipelineDB
Set cn = Nothing

write_dataLog Response.Status,"contract_selection.asp","contract_selection for the contract " &Session.contents("contrato"),Session.contents("name"),conexion ,"N/A","null","Consulta","N/A"
'---------------------------------------------------
	
	'Display Total
	If (J Mod 2) = 0 Then
		OpenTr "class=todd"
	Else
		OpenTr "class=teven"
	End If
		OpenTd "tfooter", "colspan='5' align='right'"
			Response.Write "Total bajo gestión:"
		CloseTd
		OpenTd "tfooter", "align='right'"
			Response.Write formatcurrency(Total,2)
		CloseTd

		'Display Table footer and closing tags

' By aevd  parea SF 22/02/2010
' Para consulta de extractos

'Response.write( Application("URLCertificado") & "<br>  Extractos" ) 
'response.end



	if (Session.Contents("SiteRetuns") = "2") then 
			''generar extracto   			

			If autorizarMn(2,33) Then	
			OpenTr "class=todd"		
				'OpenTd "", "align=center"			
				
				
						OpenForm "transaction", "post", Application("URLStatement") & "?TypeProcess=Extracto"&"&client=" & + ClientId + "_" + DocType , ""
						    'PlaceInput "ClientId", "hidden", ClientId, ""
							'PlaceInput "DocType", "hidden", DocType, ""		
							PlaceInput "selectextracto", "submit", "Generar Extracto", "class=sbttn"  
						CloseForm
	
						'OpenForm "menu", "post", Application("URLStatement") & "?TypeProcess=Extracto"&"&client=" & + ClientId + "_" + DocType , ""
						
						 
				''generar certificado 
				
				OpenForm "Certificado", "post", Application("URLCertificado")  & "?TypeProcess=Certificado"&"&client=" & + ClientId + "_" + DocType , ""
					'PlaceInput "ClientId", "hidden", ClientId, ""
					'PlaceInput "DocType", "hidden", DocType, ""	
					PlaceInput "selectCertificado", "submit", "Generar Certificado", "class=sbttn"  
				CloseForm
				'CloseTd
				CloseTr
	
			End If
	end if
' termina 





	CloseTr
	'Display Table footer and closing tags
	CloseTable
	CloseBody
CloseHTML
Else
	OpenTable "100%","'' align=center"
		OpenTr "class=teven"
			OpenTd "thead","align=center"
				Response.Write "No existen contratos"
			CloseTd
		CloseTr
	CloseTable
End If
OpenForm "selection", "Post", "", "target=content"
	PlaceInput "Contract", "hidden", "", ""
	PlaceInput "Product", "hidden", "", ""
	PlaceInput "Plan", "hidden", "", ""
	PlaceInput "Pas", "hidden", "", ""
	PlaceInput "Name", "hidden", "", ""
	PlaceInput "ClientId", "hidden", "", ""
	PlaceInput "DocType", "hidden", "", ""
CloseForm
'Reload Left Menu -- START
OpenForm "menu", "post", "../menu/menu.asp", "target=menu"
	PlaceInput "Name", "hidden", "", ""
	PlaceInput "ClientId", "hidden", "", ""
	PlaceInput "DocType", "hidden", "", ""
	PlaceInput "Phone", "hidden", "", ""
	PlaceInput "Contract", "hidden", "", ""
	PlaceInput "Product", "hidden", "", ""
	PlaceInput "Plan", "hidden", "", ""
	PlaceInput "Pas", "hidden", "", ""
	PlaceInput "Option", "hidden", 2, ""
CloseForm



'Reload Left Menu -- END
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>