<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		op_report.asp 9000
'Path:			/operations/radication
'Created By:		Guillermo Aristizabal 2001/09/12
'Last Modified:		A. Orozco 2001/10/08
'			Guillermo Aristizabal 2001/10/11
'                       Armando Arias - 2008/May/09 - Cumplimiento circular SF052 - PlaceTitle - write_sp_log 13290
'Parameters:
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
<!--#include file="radicationqueries.asp"-->

<%
Authorize 2,12
'Response.Write "<link rel='stylesheet' href='../../css/OLDMutualStyle1.css' type='text/css'>" & vbCrLf
'== declares ===
Dim Inicio
Dim Fin
Dim Cn
Dim rs
Dim sql
Dim I
Dim J
Dim Sel
Dim arrDetail
'== initials asignments ==
set cn = getconnpipelinedb
set rs = Server.CreateObject("ADODB.RecordSet") 
sql = "sprd_getLastRadicationNumber"
rs.Open sql,cn,3
Inicio = Request.Form.Item("Inicio")
Fin = rs.Fields(0)
rs.Close 


write_sp_log cn, 13290, "sprd_getLastRadicationNumber", 0, "", "", 0, 0, "", "op_report.asp " & _
"Loaded by " & Session("sp_miLogin")

sql = "spsp_GetSocieties"
rs.Open sql,cn,3
If rs.BOF And rs.EOF Then
	arrDetail = 0
Else
	arrDetail = rs.GetRows()
End If
rs.Close

write_dataLog Response.Status,"op_report.asp", "op_report.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"sprd_getLastRadicationNumber - spsp_GetSocieties","N/A","null","Consulta","N/A"

write_sp_log cn, 13290, "spsp_GetSocieties", 0, "", "", 0, 0, "", "op_report.asp " & _
"Loaded by " & Session("sp_miLogin")

set rs = Nothing
closeconnpipelinedb
set cn= Nothing

OpenHTML
OpenHead
        PlaceTitle "Estado Afiliacion Mfund"
%>
<link href="../../css/OLDMutualStyle.css" rel="stylesheet" type="text/css"/>
<SCRIPT LANGUAGE=javascript src=../_pipeline_scripts/validation.js></SCRIPT>
<SCRIPT LANGUAGE=javascript>

function Validate(theform){
 if (!dateValidation(theform.start_year.value,theform.start_month.value,theform.start_day.value)){     return false;
 }
 if (!dateValidation(theform.end_year.value,theform.end_month.value,theform.end_day.value)){
     return false;
 }

 var startD;
 var endD;
 startD = new Date(theform.start_year.value+ ' ' +theform.start_month.value+ ' ' +theform.start_day.value);
 endD = new Date(theform.end_year.value+ ' ' +theform.end_month.value+ ' ' +theform.end_day.value);
 if (Date.parse(endD) < Date.parse(startD) ){
   alert("Fecha Final debe ser mayor o igual a Fecha Inicial");
   return false
 }

 return true
}

function dateValidation(year_val, month_val, day_val) {
	var myDayStr = day_val;
	var myMonthStr = month_val;
	var myYearStr = year_val;
	var myDateStr = myDayStr + ' ' + myMonthStr + ' ' + myYearStr;

	/* Using form values, create a new date object
	which looks like Wed Jan 1 00:00:00 EST 1975. */
	var myDate = new Date(myDateStr);
    	
	// Convert the date to a string so we can parse it.
	var myDate_string = myDate.toGMTString();

	/* Split the string at every space and put the values into an array so,
	using the previous example, the first element in the array is Wed, the
	second element is Jan, the third element is 1, etc. */
	var myDate_array = myDate_string.split( ' ' );

	/* If we entered Feb 31, 1975 in the form, the new Date() function
	converts the value to Mar 3, 1975. Therefore, we compare the month
	in the array with the month we entered into the form. If they match,
	then the date is valid, otherwise, the date is NOT valid. */
	if ( myDate_array[2] != myMonthStr ) {
	  alert( 'La fecha ' + myDateStr + ' no es válida' );
	  return false
	}
	return true;
}

</SCRIPT>
</head>
<body class="cuerpo">
    <div class="contenido">
		</br></br>
		<div class="subtituloPagina">
			Reporte de Operaciones por Pipeline
		</div>
		</br></br>
<%

otbl"'tblcontenido'"
    opentr""
        opentd"''",""
            otbl "'tblValores'"
                OpenForm "theForm", "post", "op_rep_res.asp", "onSubmit='return Validate(this)'"			
	            OpenTr ""
		            OpenTd "'labels'", " align=right"
			            Response.Write "Fecha Inicial    "		
		            CloseTd		
		            OpenTd "'labelcombo'", "align=left"
				            Response.Write "Año:"
				            OpenCombo "start_year",  "class=listaFecha"
					            For I = 1990 To 2030
					            If I = year(date()) Then
						            Sel="selected"
					            Else
						            Sel=""
					            End If
						            PlaceItem Sel, I, I
					            Next
				            CloseCombo
				            Response.Write "Mes:"
				            OpenCombo "start_month",  "class=listaFecha"
					            PlaceItem "", "Jan", 1
					            PlaceItem "", "Feb", 2
					            PlaceItem "", "Mar", 3
					            PlaceItem "", "Apr", 4
					            PlaceItem "", "May", 5
					            PlaceItem "", "Jun", 6
					            PlaceItem "", "Jul", 7
					            PlaceItem "", "Aug", 8
					            PlaceItem "", "Sep", 9
					            PlaceItem "", "Oct", 10
					            PlaceItem "", "Nov", 11
					            PlaceItem "", "Dec", 12
				            CloseCombo
				            Response.Write "Día:"				
				            PlaceInput "start_day", "text", "01","id='RN       Día' class=listaFecha size=2 maxlength=2"
		            CloseTd
		            CloseTr

	            OpenTr ""
		            OpenTd "'labels'", "align=right"
			            Response.Write "Fecha Final"		
		            CloseTd						
		            OpenTd "'labelcombo'", "align=left"
				            Response.Write "Año:"
				            OpenCombo "end_year",  "class=listaFecha"
					            For I = 1990 To 2030
					            If I = year(date()) Then
						            Sel="selected"
					            Else
						            Sel=""
					            End If
						            PlaceItem Sel, I, I
					            Next
				            CloseCombo
				            Response.Write "Mes:"
				            OpenCombo "end_month",  "class=listaFecha"
					            PlaceItem "", "Jan", 1
					            PlaceItem "", "Feb", 2
					            PlaceItem "", "Mar", 3
					            PlaceItem "", "Apr", 4
					            PlaceItem "", "May", 5
					            PlaceItem "", "Jun", 6
					            PlaceItem "", "Jul", 7
					            PlaceItem "", "Aug", 8
					            PlaceItem "", "Sep", 9
					            PlaceItem "", "Oct", 10
					            PlaceItem "", "Nov", 11
					            PlaceItem "", "Dec", 12
				            CloseCombo
				            Response.Write "Día:"				
				            PlaceInput "end_day", "text", "01","id='RN             Día' class=listaFecha size=2 maxlength=2"				
		            CloseTd
	            CloseTr		
	            if cint(Session.Contents("sp_idSoc")) = 0 then	
		            OpenTr ""
			            OpenTd "'labels'", "align=right"
				            Response.Write "Sociedad "		
			            CloseTd						
	
			            OpenTd "'labelcombo'", "align=left"						
				            If IsArray(arrDetail) Then	
					            OpenCombo "idSociedad",  "class=listaLogin"
					            PlaceItem "", Session.Contents("sp_idSoc"),"Todas"
					            For J = 0 To UBound(arrDetail, 2)												
						            PlaceItem "", arrDetail(0,J), arrDetail(1,J)								
					            Next
					            CloseCombo
				            Else
					            PlaceInput "idSociedad","hidden",Session.Contents("sp_idSoc")	
			            CloseTd
		            CloseTr						
	            End if		
	            Else
			            PlaceInput "idSociedad","hidden",Session.Contents("sp_idSoc"),""
	            End If
	
	            OpenTr ""
		            OpenTd "''", "align=left colspan=4"						
			            Response.Write "&nbsp;"
		            CloseTd
	            CloseTr						

	            OpenTr "align = center"
		            OpenTd "''", "align=center colspan=4"				 
			            PlaceInput "Enviar","submit","Enviar","class=button-OLD" 
		            CloseTd
	            CloseTr				
              CloseForm
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

If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If

%>