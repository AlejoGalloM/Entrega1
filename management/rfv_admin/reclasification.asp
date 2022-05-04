<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../../operations/_pipeline_scripts/url_check.asp"-->
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<%
'File Name:			reclasification.asp
'Path:				management/rfv_admin
'Created By:		A. Orozco 2004/05/17
'Last Modified:	
'						
'Modifications:	
' Added AUM Field
'Parameters:		
'Returns:
'Additional Information:
'	Uses [TRFV].dbo.GetReclasificationClients

On Error Resume Next
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1

Dim adoConn
Dim Sql
Dim rstReclasification
Dim arrClients
Dim arrRow, arrCol
Dim I, RecCount
Dim Pages, PgCount
' ***************** POR HACER ****
' Authorize 0,8			AUTORIZACION !!!!!!!!!!!
' *****************

Set adoConn = GetConnPipelineDB

Sql = "[TRFV].dbo.GetReclasificationClients"
Set rstReclasification = Server.CreateObject("ADODB.Recordset")
rstReclasification.PageSize = Application("PagesHistory")
rstReclasification.Open Sql, adoConn

write_dataLog Response.Status,"reclasification.asp", "reclasification.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"[TRFV].dbo.GetReclasificationClients","N/A","null","Consulta","N/A"

Pages = rstReclasification.PageCount

If rstReclasification.BOF And rstReclasification.EOF Then
	arrClients = 0
Else
	arrClients = rstReclasification.GetRows()
End If

' ***************** POR HACER ****
'	DETERMINAR EL LOG !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'	write_sp_log adoConn, 8700, "", 0, "", "", 0, 0, "", "management/default.asp Loaded by " & Session("sp_miLogin")
' ***************** 

CloseConnPipelineDB
Set adoConn = Nothing

OpenHTML
OpenHead
PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
PlaceMeta "Pragma", "", "no_cache"
PlaceLink "REL", "stylesheet", "../../css/OLDMutualStyle.css", "text/css"
%>
<script LANGUAGE="javascript" src="../../operations/_pipeline_scripts/validation.js"></SCRIPT>
<script language="javascript">
	function CheckAll(req)
	{
		form = document.reclasif
		with (form)
		{
			for (i = 0; i <= elements.length - 1; i ++)
			{
				if (elements[i].type == 'checkbox')
				{
					if (req.checked)
					{
						elements[i].checked = true
					}
					else
					{
						elements[i].checked = false
					}
				}
			}
		}
	}
</script>
</head>
<body class="cuerpo">
	<div class="contenido">
		</br></br>
		<div class="subtituloPagina">
            Reclasificación Clientes
		</div>
		</br></br>
<%


otbl"tblcontenido"
    opentr""
        opentd"",""
            otbl"tblvalores"
	            OpenTr "valign=middle"
		            OpenTd "''", "align=center valign=middle"
			            otbl"tblvalores"
				                If  Not (rstReclasification.BOF And rstReclasification.EOF) Then
					                If Request.Form("Page") = "" Then
						                PgCount = 1
					                Else
						                PgCount = Request.Form("Page")
					                End If
					                rstReclasification.AbsolutePage = PgCount
					                '** Page Number Buttons
					                OpenTr ""
					                    OpenForm "pagesTop", "post", "", ""
						                    OpenTd "''", ""
							                    For I = 1 To Pages
								                    PlaceInput "page" & I, "button", I, "class=button-OLD onClick='javascript:pagesTop.Page.value=" & _
								                    I & "; pagesTop.submit()'"
							                    Next
							                    PlaceInput "Page", "hidden", "", ""
						                    CloseTd
					                    CloseForm 'pagesTop
					                CloseTr
					                OpenTr ""
						                '** Previous page button
						                OpenForm "prev", "post", "", ""
						                    OpenTd "''", ""
							                    If  PgCount > 1 Then
									                    PlaceInput "submit","submit", "< Anterior", "class=button-OLD"
									                    PlaceInput "Page", "hidden", PgCount - 1, ""
							                    Else
								                    Response.Write "&nbsp;"
							                    End If
						                    CloseTd
						                CloseForm 'prev
						                '** Page Count
						                OpenTd "''", ""
							                Response.Write "Pagina: " & PgCount & " de " & Pages
							                Response.Write " - No. Total de registros: " & rstReclasification.RecordCount
						                CloseTd
						                '** Next Page Button
						                OpenForm "next", "post", "", ""
						                    OpenTd "''", ""
							                    If  rstReclasification.AbsolutePage < Pages Then
									                    PlaceInput "submit","submit", "Siguiente >", "class=button-OLD"
									                    PlaceInput "Page", "hidden", PgCount + 1, ""
							                    Else
								                    Response.Write "&nbsp;"
							                    End If
						                    CloseTd
						                CloseForm 'next
					                CloseTr
				                '****** CLIENTS TO RECLASIFY FORM
				                OpenForm "reclasif", "post", "reclasificationProcess.asp", "onSubmit='javascript:return formValidation(this)'"
				                ' Table titles...
				                    OpenTr ""
					                    OpenTh "''", "align=center"
						                    PlaceInput "ChkAll", "checkbox", "", "onClick='CheckAll(this)'"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Documento"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Tipo Documento"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Nombre"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Financial Planner"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Sociedad"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Ultimo Segmento Asignado"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Ultimo Segmento Calculado"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "Característica Ultimo Segmento Calculado"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "AUM Período Calculado"
					                    CloseTh
					                    OpenTh "''", "align=center"
						                    Response.Write "No. Periodos Bajando"
					                    CloseTh
				                    CloseTr
				                '*** Table rows
					                RecCount = 0
					                Do While Not rstReclasification.EOF And RecCount < rstReclasification.PageSize
						                ' Check alternating items
						                If (rstReclasification.AbsolutePosition Mod 2) = 0 Then
							                OpenTr ""
						                Else
							                OpenTr ""
						                End If
						                ' Display Item's checkbox
						                        OpenTd "''", ""
							                        PlaceInput "Client_"& RecCount, "checkbox", _
							                        Trim(rstReclasification(0)) & "|" & Trim(rstReclasification(1)) & "|" & Trim(rstReclasification(2)), ""
						                        CloseTd
						                    ' Display each column in the row
						                        For arrCol = 1 To rstReclasification.Fields.Count - 1
							                        OpenTd "''", "align=center"
								                        If IsNull(rstReclasification(arrCol)) Then
									                        Response.Write "&nbsp;"
								                        Else	
									                        Response.Write rstReclasification(arrCol)
								                        End If
							                        CloseTd
						                        Next
						                    CloseTr
						                RecCount = RecCount + 1
						                rstReclasification.MoveNext
					                Loop
					                rstReclasification.MoveLast
				                CloseForm 'reclasif
					            ' Display bottom navigation buttons
					            OpenForm "pagesBottom", "post", "", ""
						            OpenTd "''", ""
							            For I = 1 To Pages
								            PlaceInput "page" & I, "button", I, "class=button-OLD onClick='javascript:pagesBottom.Page.value=" & _
								            I & "; pagesBottom.submit()'"
							            Next
							            PlaceInput "Page", "hidden", "", ""
						            CloseTd
					            CloseForm 'pagesBottom
					            OpenTr ""
					                OpenForm "prev", "post", "", ""
					                    OpenTd "''", ""
						                    If  CInt(PgCount) > 1 Then
								                    PlaceInput "submit","submit", "< Anterior", "class=button-OLD"
								                    PlaceInput "Page", "hidden", PgCount - 1, ""
						                    Else
							                    Response.Write "&nbsp;"
						                    End If
					                    CloseTd
					                CloseForm 'prev
					                OpenTd "''", ""
						                Response.Write "Pagina: " & PgCount & " de " & Pages
						                Response.Write " - No. Total de registros: " & rstReclasification.RecordCount
					                CloseTd
					                OpenForm "next", "post", "", ""
					                        OpenTd "''", ""
						                        If  CInt(PgCount) < CInt(Pages) Then
								                        PlaceInput "submit","submit", "Siguiente >", "class=button-OLD"
								                        PlaceInput "Page", "hidden", PgCount + 1, ""
						                        Else
							                        Response.Write "&nbsp;"
						                        End If
					                        CloseTd
					                    CloseForm 'next
					            CloseTr
				            Else
					            OpenTr ""
						           OpenTd "'texto-informativo'", ""
							           Response.Write "No Hay Registros"
						           CloseTd
					           CloseTr
				            End If
			            ctbl
		            CloseTd
	            CloseTr
	            OpenTr "valign=middle"
		            OpenTd "''", "align=center valign=middle colspan=2"
			            Response.Write "&nbsp;"
		            CloseTd
	            CloseTr
	            If IsArray(arrClients) Then
		            OpenTr ""
			            OpenTd "''", "align=center valign=middle colspan=2"
				            PlaceInput "Send", "Button", "Reclasificar", "class=button-OLD onClick='document.reclasif.submit()'"
				            PlaceInput "Send", "Button", "Volver", "class=button-OLD onClick='document.returnForm.submit()'"
			            CloseTd
		            CloseTr
	            Else
		            OpenTr ""
			            OpenTd "''", "align=center valign=middle colspan=2"
				            PlaceInput "Send", "Button", "Volver", "class=button-OLD onClick='document.returnForm.submit()'"
			            CloseTd
		            CloseTr
	            End If
            ctbl
        closetd
    closetr
ctbl

%>
        <p/>
        <p/>
		</div>
	</body>
</html>
<%
OpenForm "menu", "post", "../../operations/menu/menu.asp", "target=menu"
	PlaceInput "Product", "hidden", "Gestion",	""
	PlaceInput "Option", "hidden", 3, ""
CloseForm
If Err.Number <> 0 Then
	Dim url
	url = Application("ErrorURL") & "?ErrNum=" & Err.number & "&ErrSource=" & Err.Source & "&ErrDesc=" & Err.Description & "&page=reclasification.asp"
	url = URLEncode(url)
	Response.Redirect url
End If
%>
<Form name=returnForm action=gestionsegmentacion.asp method=post>
</Form>