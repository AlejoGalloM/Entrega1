<%@ Language=VBScript %>
<%
'===================================================================================
'@author name:		 		J carreno 
'@exception name:				
'@param name description:	date 
'@return					
'@since						2002/05/17
'@version					1.0
'@File Name:				disburCriteria.asp [13208]
'@Path:						insurance/admon
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

Authorize 4,17

Dim objConn
dim i
dim strsql
dim combo
dim cadenacombo1


'=========================================================
'==create objects
'=========================================================
Set objConn = GetConnPipelineDB
'=========================================================
'==write log
'=========================================================
write_sp_log objConn, 13208, "", 0, "", "", 0, 0, "", "disburcriteria.asp insurance Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"disburcriteria.asp", "disburcriteria.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"","N/A","null","Consulta","N/A"

CloseConnPipelineDB

'get list month
cadenacombo1="<select name=mes>"
for i=1 to 12 
		cadenacombo1=cadenacombo1& "<option value="&i&">"&i&"</option>" 	
next
cadenacombo1=cadenacombo1&"</select>"

OpenHTML
	OpenHead
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
        response.Write "<SCRIPT LANGUAGE=javascript src='../js/scripts.js'></SCRIPT>" & chr(10)
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
		PlaceLink "REL", "stylesheet", "../../css/style1.css", "text/css"
%>
<script language="javascript">


  function validar(){

    if (!positivo(theform.ano, theform.ano.value,0)) return false;
    mydate=new Date();
	if (theform.ano.value*1>mydate.getYear()){
      alert('ano desde invalido ');
      theform.ano.select();theform.ano.focus();
      return false;	
	}
    return true;
  }


</script>
<%

	CloseHead
	OpenBody "''", "onLoad='document.theform.ano.focus()'"
OpenTable "75%", "'' align=center"
	OpenTr ""
		OpenTd "thead", "colspan=4"
		Response.Write "<br>" & vbCrLf		
			Response.Write "Primas No Cobradas" & vbCrLf
			Response.Write "<br><br>" & vbCrLf
		CloseTd
	CloseTr
	OpenTr ""
		OpenTd "teven", "colspan=4"
			Response.Write "&nbsp;"
		CloseTd					  
	CloseTr
	OpenForm "theform", "post", "disbursement.asp", " onSubmit='return validar()'"
	  placeinput "opcion","hidden",Request.Form("opcion"),""
	OpenTr ""
		OpenTd "tbody", ""
			OpenTable "''", "'' align=center "
					OpenTr ""
						OpenTd "tbody", ""
							response.Write " Año: "				
						CloseTd
						OpenTd "tbody", ""
							placeinput "ano","text",datepart("yyyy",now())," size=5 "
						CloseTd
					closetr
					opentr ""	
						OpenTd "tbody", ""
							response.Write " Mes: "
						CloseTd
						OpenTd "tbody", ""
							response.Write cadenacombo1				
						CloseTd
					closetr
			CloseTable
			
		CloseTd
		OpenTd "", ""
			Response.Write "&nbsp;"
		CloseTd					  				
	CloseTr



	OpenTr ""
		OpenTd "teven", "colspan=4"
			Response.Write "&nbsp;"
		CloseTd					  
	CloseTr	

	OpenTable "''", "'' align=center "
			OpenTr ""
				OpenTd "thead", "align=left nowrap"
					PlaceInput "btnDate", "submit", "Imprimir", "id='               Enviar' class=sbttn2"
				CloseTd
				OpenTd "tbody", ""
					placeinput "botton","button","Regresar"," onclick=window.location='centraladmon.asp' class='sbttn2' "
				CloseTd
			CloseTr
	CloseTable

CloseTable

	CloseForm

Response.Write "<p></p><p></p>"

	CloseBody
CloseHTML
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>