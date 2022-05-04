<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		contract_irr_by_date.asp 500
'Path:				contract_info/
'Created By:		Alejandro Pulgarin Correa
'Parameters:		User must be logged on
'						Session("docNumber")
'Returns:			MFUND contract IRR information
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
<%
Authorize 7,2
Dim Sql 'SQL Sentences holder
Dim objConn 'Database Connection
Dim rstDescription 'Results recordset
Dim objRst	' Strategist PAS recordset
Dim Sel
Dim rs, cn
Dim Param,obConnPlus
Dim I,J 'Asset & Standing Alloc counters
Dim Contract, Product, Plan, ClientId, Name, Phone, OU, Pas, PasName, Frase
dim Reference '<I&T - DMPC>

Contract = Request.Form("Contract")
Product = Trim(Request.Form("Product"))
Plan = Request.Form("Plan")
ClientId = Request.Form("ClientId")
Name = Request.Form("Name")


Set objConn = GetConnPipelineDB

set objRst= Server.CreateObject("ADODB.Recordset")
set objConn=GetConnpipelineDB

write_sp_log objConn, 8800, "", Contract, Product, "", ClientId, 0, "", "contract_irr.asp loaded by " & _
Session.Contents("sp_milogin")

write_dataLog Response.Status,"contract_irr_by_date.asp","contract_irr_by_date.asp " &"- " & Session("sp_miLogin"),Session.contents("name"),"null" ,"N/A","null","Consulta","N/A"

'==========================================================================
''<I&T - DMPC 2009/02/10 - Modificado por proceso de Referencia Unica - Recaudos>
' Obtiene la referencia única a partir del Contrato y el producto
'==========================================================================
 Reference = GetReferenciaUnica( Product, Contract)
'==========================================================================

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
	CloseHead
	PlaceTitle "Rentabilidad del contrato"
	OpenBody "", ""

%>
<script language=javascript> 
function validate(form) 
{ 
	var value = false;
	with (form) 
	{ 
		sendForm(form);
		value = dateValidation(s_yval.value, s_mval.value, s_dval.value); 
		value = value && dateValidation(e_yval.value, e_mval.value, e_dval.value); 
	} 

	if (!value) 
	{ 
		enableButtons(form);
	} 
	return value;
} 
</SCRIPT>
<%

		Response.Write "<br><br>"
		OpenTable "70%","border=0"
			OpenTable "80%", "t_table align=center"
				OpenTr "class=teven"
					OpenTd "thead", "align=center colspan=3"
						Response.Write "Reporte de Rentabilidad"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "", "align=center colspan=3"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr	

				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "<STRONG>Producto</STRONG>"
					CloseTd
					OpenTd "tbody", ""
						Response.Write Product
					CloseTd
				CloseTr	
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "<STRONG>Contrato</STRONG>"
					CloseTd
					OpenTd "tbody", ""
						Response.Write Reference 
					CloseTd
				CloseTr	
				OpenTr ""
					OpenTd "tbody", ""
						Response.Write "<STRONG>Nombre</STRONG> "
					CloseTd
					OpenTd "tbody", ""
						Response.Write Name
					CloseTd
				CloseTr	

				OpenTr ""
					OpenTd "", "align=center colspan=3"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr	
				
				OpenTr "class=teven"
					OpenTd "thead", "align=center colspan=3"
						Response.Write "Consulta a fecha pasada"
					CloseTd
				CloseTr
				OpenTr ""
					OpenTd "", "align=center colspan=3"
						Response.Write "&nbsp;"
					CloseTd
				CloseTr	
				OpenTr ""
					OpenTd "tbody", "align=center width=10%"
						Response.Write "&nbsp;"
					CloseTd

					OpenTd "thead", "align=right width=20%"
						Response.Write "Fecha a Consultar&nbsp;&nbsp;"
					CloseTd
					OpenForm "sel_irr_date", "post", "contract_irr_by_date_process.asp", ""
						OpenTd "tbody", "align=left width=40%"
							Response.Write "Año: "
							OpenCombo "e_yval", "class=bttntext"
								For I = 1990 To 2030
									If Year(Date) = I Then
										Sel = "selected"
									Else
										Sel = ""
									End If
									PlaceItem sel, I, I
								Next
							CloseCombo
							Response.Write "Mes: "
							OpenCombo "e_mval", "class=bttntext"
								For I = 1 To 12
									If Month(Date) = I Then
										Sel = "selected"
									Else
										Sel = ""
									End If
									PlaceItem sel, MonthName(I, 1), I
								Next
							CloseCombo
							Response.Write "Día: "
							OpenCombo "e_dval", "class=bttntext"
								For I = 1 To 31
									If Day(Date) = I Then
										Sel = "selected"
									Else
										Sel = ""
									End If
									PlaceItem sel, I, I
								Next
							CloseCombo
						CloseTd
					CloseTr
					OpenTr ""
						OpenTd "tbody", "align=center colspan=3"
							PlaceInput "go", "submit", "Continuar", "class=sbttn"
						CloseTd
					CloseTr
					PlaceInput "Name", "hidden", Replace(Name," ","%20"), ""
					PlaceInput "Contract", "hidden", Contract, ""
					PlaceInput "Product", "hidden", Product, ""
				CloseForm
			CloseTable
	CloseBody
CloseHTML

If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>
