<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		operations.asp
'Path:				/
'Created By:		A. Orozco 2001/07/24
'Last Modified:			J. Carreño 2002/09/25, add security context
'						A. Orozco 2001/09/10
'						A. Orozco 2001/10/08
'						Guillermo Aristizabal 2001/10/11
'Parameters:		none
'Returns:			
'Additional Information: Operations home page
'===================================================================================
'Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="_pipeline_scripts/tags.asp"-->
<!--#include file="_pipeline_scripts/pipeline_scripts.asp"-->
	<script language="javascript">
        var now = new Date();
		now.setTime(now.getTime());
		var accessDate = now.getFullYear() + "-" + (now.getMonth() + 1) + "-" + now.getDate() + " T " + now.getHours() + ":" + now.getMinutes() + ":" + now.getSeconds();
		document.cookie = "accessDate=" + accessDate + ";expires=" + "; path=/super_pipeline";
    </script>
<%
Dim objConn

Set objConn = GetConnPipelineDB
write_sp_log objConn, 2, "", 0, "", "", 0, 0, "", "operations.asp Loaded by " & Session("sp_miLogin")
CloseConnPipelineDB
Set objConn = Nothing

'=====begin modification 2002/09/25
if len(trim(Request.Form("initial_page"))) = 0 then
	Response.redirect Application("ErrorURL")
end if
'=====end modification 2002/09/25

Session("browser")=Request.Cookies("browser")
Session("idworker")=Request.Cookies("sp%5Fidworker")
Session("ipPublica")=Request.Cookies("ipPublica")
Session("accessDate")=Request.Cookies("accessDate")

write_dataLog Response.Status,"operations.asp","operations.asp Loaded by " & Session("sp_miLogin"),"null","null","null","null","Login","N/A"

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV='expires' content='Wednesday, 27-Dec-95 05:29:10 GMT'>
<META HTTP-EQUIV='Pragma' CONTENT='no_cache'>
<title>Pipeline 2.0 - Colombia</title>
</head>
<frameset rows="*" frameborder="NO" border="0" framespacing="0" cols="20%,80%" onLoad> 
	<frame name="menu" noresize scrolling=auto frameborder="0" marginwidth="0" marginheight="0" src="menu/menu.asp" border=0 framespacing=0>
	<frame name="content" noresize scrolling="auto" frameborder="0" marginwidth="0" marginheight="0" src="<%=Request.Form("initial_page")%>">
</frameset>
<noframes> 
	<body bgcolor="#FFFFFF" text="#000000">
	</body>
</noframes> 
</html>
