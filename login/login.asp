<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		login.asp  0
'Path:				/login
'Created By:		Andres Felipe Orozco 2001/07/16
'Last Modified:	G. Aristizabal  2001/07/18 add the select initial_page
'						A. Orozco 2001/09/03
'						Guillermo Aristizabal  2001/09/18 auth & log
'						A. Orozco 2001/09/24 Site Availability function
'						A. Orozco 2001/10/11
'						A. Orozco 2002/04/03 Added AGENTES to the combo box
'						A. Orozco 2002/04/23 Added AGENTES Application variable
'						J. carreno 2002/05/04 Added SEGUROS to the combo box
'						A. Orozco 2002/08/02 Added Agents Logout redirection.
'						D. Pérez 2008/05/06 Added PlaceTitle "Login" and changed in write_sp_log the parameter page_id=13105
'						A. Alarcon 2013/07/12 Added div tags for page distribution and changed the logo
'	Julian Zapata Desarrollo-IT 02-02-2017 Se modifica el Texto para que aparezca Financial Planner / Asesor de Seguros
'Parameters:		none
'Returns:			
'Additional Information: Login page for Pipeline
'===================================================================================
Option Explicit
On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../operations/_pipeline_scripts/pipeline_scripts.asp"-->

<script language="javascript">
    var browser;
    if (navigator.userAgent.toLowerCase().indexOf('edg') > -1) {
        browser= "Edge" ;
    }
    else {
        var getBrowserInfo = function () {
            var ua = navigator.userAgent, tem,
                M = ua.match(/(opera|chrome|safari|firefox|msie|trident(?=\/))\/?\s*(\d+)/i) || [];
            if (/trident/i.test(M[1])) {
                tem = /\brv[ :]+(\d+)/g.exec(ua) || [];
                return 'IE ' + (tem[1] || '');
            }
            if (M[1] === 'Chrome') {
                tem = ua.match(/\b(OPR|Edge)\/(\d+)/);
                if (tem != null) return tem.slice(1).join(' ').replace('OPR', 'Opera');
            }
            M = M[2] ? [M[1], M[2]] : [navigator.appName, navigator.appVersion, '-?'];
            if ((tem = ua.match(/version\/(\d+)/i)) != null) M.splice(1, 1, tem[1]);
            return M.join(' ');
        };
        browser = getBrowserInfo();
    }
	document.cookie = "browser=" + browser + ";expires=" + "; path=/super_pipeline";

	const request = new XMLHttpRequest()
	var resultado;
    request.onreadystatechange = () => {
        if (request.readyState === 4) {
            if (request.status === 200) {
				resultado = request.responseText;
                document.cookie = "ipPublica=" + resultado + ";expires=" + "; path=/super_pipeline";
            }
        }
    }
    request.open("GET", "https://api.ipify.org/")
	request.send()
    
</script>


<%
Dim strSql 'Stored procedures and SQL queries
Dim objConn 'ADODB Connection
Dim objRst 'ADODB Recordset
Dim Avail, Msg 'Get the site status
Dim dayDate,timeDate,exitDate

dayDate= Date
timeDate= Time
exitDate= year(dayDate)&"-"&month(dayDate)&"-"&Day(dayDate)&" "&timeDate

Avail = Available
Select Case Avail
	Case 0 'Available
		Msg = ""
	Case 1 'Not Available
		Msg = Application("Not_Available")
	Case 2 'Partially Available
		Msg = Application("Part_Available")
End Select
Set objConn = GetConnpipelineDB
write_sp_log objConn, 13105, "", 0, "", "", 0, 0, "", "login.asp loaded"
If Request.QueryString("out") = 1 Then 
	Session.Contents.Item("sp_flag")="out"  
	write_sp_log objConn, 13105, "", 0, "", "", 0, 0, "", "Session Ended"
	if Session.contents("accessDate")<>"" then
	write_dataLog Response.Status,"login.asp","Session Ended by " & Session("sp_miLogin"),"null","null","prueba","null","Session Ended",exitDate
	end if
	Session.Abandon()	
End if
' esto es de aev para portal distribuidor
if int(mid(Request.ServerVariables("REMOTE_ADDR"),1,3))<>10 then 
	write_sp_log objConn, 13105, "", 0, "", "", 0, 0, "", "Redirecionado al portal distribuidores"
	Response.Redirect(Application("Dis_URL")) 
End if
'termina 
if Session.Contents("sp_flag") = "in"  Then
	write_sp_log objConn, 13105, "", 0, "", "", 0, 0, "", "User already logged in"
	Response.Redirect("relogin.asp")
End if
strSql = "exec sppl_lockip '" & Request.ServerVariables("REMOTE_ADDR") & "'"
Set objRst = Server.CreateObject("ADODB.Recordset")
objRst.Open strSql, objConn
If objRst.Fields(0) > 49 Then
	CloseConnpipelineDB
	write_sp_log objConn, 13105, "", 0, "", "", 0, 0, "", "Users IP is locked"
	Response.Redirect "Sorry_page.asp?error=3"
End If
CloseConnpipelineDB
%>
<html>
    <head>
        <title>Old Mutual - Pipeline</title>
        <link href="../css/OLDMutualStyle.css" rel="stylesheet" type="text/css"/>
        <meta name="" http-equiv="expires" content="Wednesday, 27-Dec-95 05:29:10 GMT"/>
        <meta name="" http-equiv="Pragma" content="no_cache"/>
        <script type="text/javascript" language="javascript" src='../operations/_pipeline_scripts/validation.js'></script>
    </head>
    <%
	If Avail = 1 Then
    %>
    <body class="loginbody">
    <%
    Else
    %>
    <body class="loginbody" onload="document.login.login.focus()">
    <%	
    End If
%>
	<div id="completo">
		<div id="contenedor">
			<table width="100%" height="81px" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr>
					<td height="81px">
						<table width="100%" border="0" cellspacing="0" cellpadding="0" height="81px">
							<tr>
								<td height="3px"/>
							</tr>
							<tr>
								<td id = "encabezado">
								</td>
							</tr>
							<tr>
								<td id="logo-head" class="logotipo">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<div id="logo-producto">
				<div id="logo-productoizq">
				</div>
				<div id="logo-productoder">
				</div>
			</div>			
			<div id="cuerpo">
				<table class="tblPrincipal"> 
					<tr>
						<td class="table_cont">
							<div id="cell-izquierda">
								<div id="leyenda">
									<font size="+0">
												<p class="txlogon">Usted acaba de ingresar al
												sitio que Skandia ha provisto para que usted pueda
												realizar sus operaciones como Financial
												Planner / Asesor de Seguros.</p>
												<p class="txlogon">Ahora debe proceder a digitar su login 
												y password interno como usuario de 
												Pipeline.</p>
									</font>
									<br>
										<a href="<%=Application("SalesForceLink")%>" target='_blank' class="enlace">Acceso a Sales Force
								</div>
							</div>
						</td>
						<td>
							<div id="login">
							<%
								If Avail <> 1 Then
									OpenForm "login", "post", "logon.asp", "onSubmit='return formValidation(this)' autocomplete='off'"
										OpenTable "''", "tbllogin"
											OpenTr ""
												OpenTd "'table_login'", ""
												CloseTd
												OpenTd "'table_login1'", ""
												CloseTd
											CloseTr
											OpenTr ""
												OpenTd "'txfields table_login alturaLogin'", "align=right"
													Response.Write "<b>Login</b>"
												CloseTd
												OpenTd "'table_login1'", ""
													PlaceInput "login", "text", "", "class=fieldsLogin id='R              Login'"
												CloseTd																																										
											CloseTr
											OpenTr ""
												OpenTd "'txfields table_login alturaLogin'", "align=right"
													Response.Write "<b>Password</b>"
												CloseTd
												OpenTd "'table_login1'", ""
													PlaceInput "password", "password", "", "class=fieldsLogin id='R              Password'"
												CloseTd
											CloseTr
											OpenTr ""
												OpenTd "'txfields table_login alturaLogin'", "align=right"
													Response.Write "<b>Inicio</b>"
												CloseTd
												OpenTd "'table_login1'", ""
													OpenCombo "initial_page", "class=listaLogin"
													PlaceItem "", "../admin/pipeline_admin/default.asp","Administración"
													PlaceItem "", Application("agents_url"), "Agentes"
													PlaceItem "", "../operations/radication/default.asp","Documentos"
													PlaceItem "", "../management/default.asp","Gestión"
													PlaceItem "", "../info/default.asp", "Información"
													PlaceItem "selected", "../operations/search/search.asp","Operaciones"
													PlaceItem "", "../insurance/insurance/default.asp","Seguros"
													PlaceItem "", "../skandia_university/default.asp","Skandia University"
													CloseCombo
												CloseTd
											CloseTr
											OpenTr ""
												OpenTd "'table_login'", "align=right"
												CloseTd
												OpenTd "'table_login1'", "align=left"
													PlaceInput "btnSend", "submit", "Enviar", "id='               Enviar' class=button-OLD"
													PlaceInput "close", "button", "Cerrar", "class=button-OLD onclick=javascript:window.close()"
												CloseTd
											CloseTr
											OpenTr ""
												OpenTd "'warning table_login1'", "align=center colSpan=2"
													Response.Write "<br><p>" & Session.Contents("sp_error") & "</p>" 
												CloseTd												
											CloseTr
												OpenTr ""
                                            	OpenTd "'table_login'", ""
												CloseTd
													OpenTd "'warning table_login1'", "align=center"
														Response.Write "<br><p>" & Msg & "</p>" 
													CloseTd														
												CloseTr
										CloseTable
									CloseForm								
							%>
							<%Else%>
								<div class="warning table_login1" align="center">
									<%
									Response.Write "<br><p>" & Msg & "</p>"
									%>
								</div>
							<%End IF%>
							</div>
						</td>
					</tr>
				</table>
			</div>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td id = "pie">
					</td>
				</tr>
			</table>
		</div>
	</div>
    </body>
</html>
<%
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>