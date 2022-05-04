<%@ Language=VBScript %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../../operations/_pipeline_scripts/url_check.asp"-->
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<%
'===================================================================================
'File Name:		adminsmlv.asp xxyy
'Path:				management/rfv_admin
'Created By:		G. Pinerez 2002/02/26
'Last Modified:	
'						
'Modifications:	
'Parameters:		none
'
'Returns:			Mangement default page
'Additional Information:
'ToDoes : "POR HACER ****"
'===================================================================================
On Error Resume Next
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
Dim adoConn
Dim strSQL
Dim arrParameters
Dim objRst 'Recordset object
Dim Modificando
Dim Clasificacion
Dim Smlv
Dim processInfo, component_id
Modificando = Request.Form("Modificando")
Smlv = Request.Form("Smlv")
' ***************** POR HACER ****
' Authorize 0,8			AUTORIZACION !!!!!!!!!!!
' *****************
Set adoConn = GetConnPipelineDB
' ***************** POR HACER ****
'	DETERMINAR EL LOG !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'	write_sp_log adoConn, 8700, "", 0, "", "", 0, 0, "", "management/default.asp Loaded by " & Session("sp_miLogin")
' ***************** 
'Set adoConn = Server.CreateObject("ADODB.Connection")
if Modificando = "True" then
	smlv = Replace (smlv,",",".")
	adoConn.Execute "exec trfv..sprfv_addsmlvParameter " & (Smlv)
end if
Set objRst = Server.CreateObject("ADODB.Recordset")

component_id = "adminsmlv.asp"
processInfo =  "management/default.asp Loaded by " & Session("sp_miLogin")

'on error goto 0 
strSQL = "exec trfv.dbo.sprfv_getsmlvParameter"
objRst.Open strSQL, adoConn, 3
Session("smlvOld")=arrParameters(0,0)
arrParameters = objRst.GetRows()
smlv = arrParameters(0,0)
objRst.Close
CloseConnPipelineDB
Set adoConn = Nothing

OpenHTML
OpenHead
PlaceMeta "Pragma", "", "no_cache"
PlaceLink "REL", "stylesheet", "../../css/OLDMutualStyle.css", "text/css"
%>
<script language="javascript" src="../../operations/_pipeline_scripts/validation.js"></script>
</head>
<body class="cuerpo">
    <div class="encabezado">
         Pipeline
    </div>
    <div class="rounded">
        <b class="r1"></b><b class="r2"></b><b class="r3"></b><b class="r4"></b>
    </div>
	<div class="contenido">
		<div class="subtituloPagina">
            Salario M�nimo Legal Vigente 
		</div>
<%


otbl"tblcontenido"
    opentr""
        opentd"",""
            otbl"tblvalores"
                    opentr""
                        opentd"",""
                            Response.Write "<br><br>"
                        closetd
                    closetr
		            OpenTr ""			
			            OpenTd "'titulotabla'", "align=center colspan=2 "
				            Response.Write " Salario m�nimo legal vigente " & _ 
                            FormatCurrency( smlv ,2 )
			            CloseTd
		            CloseTr
		            OpenTr ""
		            OpenForm "consi", "post", "adminsmlv.asp", "onSubmit='javascript:return formValidation(this)'"
			            OpenTd "''", "align=center"				
				            PlaceInput "Smlv", "text", arrParameters(0,0) , " class=txboxGenericas"
			            CloseTd
			            OpenTd "''", "align=center "
				            PlaceInput "Modificando", "hidden", "True", ""
				            PlaceInput "Adicionar", "submit", "Adicionar", "class=button-OLD"
			            CloseTd
		            CloseForm
		            CloseTr
            ctbl
        closetd
    closetr
    opentr""
        opentd"",""
            Response.Write "<br><br>"
        closetd
    closetr
    OpenTr ""
	    OpenTd "", "colspan=2 "	
			write_dataLog Response.Status,component_id,processInfo,Session.contents("idworker"), "sprfv_addsmlvParameter - trfv.dbo.sprfv_getsmlvParameter" ,Session.contents("smlvOld"),FormatCurrency( smlv ,2 ),"Operación-Modificación","N/A"
		    
			OpenForm "cons", "post", "administraciontrfv.asp", "onSubmit='javascript:return formValidation(this)'"			
		    OpenForm "cons", "post", "administraciontrfv.asp", "onSubmit='javascript:return formValidation(this)'"
		        PlaceInput "Retornar", "submit", "Retornar", "class=button-OLD"
		    CloseForm
	    CloseTd				
    CloseTr
ctbl

%>
	<p/>
        <p/>
		</div>

		<div class="rounded">
			<b class="r4"></b><b class="r3"></b><b class="r2"></b><b class="r1"></b>
		</div>		
	</body>
</html>

<%

%>
