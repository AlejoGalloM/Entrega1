<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:					rad_rep_withdrawal.asp 9100
'Path:							/operations/radication
'Created By:					Guillermo Aristizabal 2001/09/03
'Last Modified:				A. Orozco 2001/09/19
'									A. Orozco 2001/10/08
'				Guillermo Aristizabal 2001/10/11
'Modifications:				
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
Authorize 3,12
Response.Write "<link rel='stylesheet' href='../../css/OLDMutualStyle.css' type='text/css'>" & vbCrLf
''== declares ===
Dim I
Dim Sel
Dim cn

Set cn = GetConnPipelineDB

write_sp_log cn, 9100, "", 0, "", "", 0, 0, "", "rad_rep_withdrawal.asp " & _
"Loaded by " & Session("sp_miLogin")

write_dataLog Response.Status,"rad_rep_withdrawal.asp", "rad_rep_withdrawal.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"","N/A","null","Consulta","N/A"

CloseConnPipelineDB

OpenHTML
OpenHead
%>
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
	//var myDate = new Date(year_val, month_val, day_val);
	
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
            Reporte de Radicación de Retiros
		</div>
		</br></br>
<%

otbl"tblcontenido"
    opentr""
        opentd"''",""
            otbl"tblvalores"
                OpenForm "theForm", "post", "rad_rep_wd_res.asp", "onSubmit='return Validate(this)'"
                    opentr""
                        opentd"",""
                            Response.Write "<br><br>"
                        closetd
                    closetr			
	                OpenTr ""
		                OpenTd "'labels'", "align=left"
			                Response.Write "Fecha Inicial "		
		                CloseTd						
		                OpenTd "'labelcombo'", ""
				                Response.Write "Año:"
				                OpenCombo "start_year",  "class=listafecha"
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
				                OpenCombo "start_month",  "class=listafecha"
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
				                PlaceInput "start_day", "text", "01","id='RN             Día' class=listafecha size=2 maxlength=2"
		                CloseTd
		            CloseTr
	                OpenTr ""
		                OpenTd "'labels'", "align=left"
			                Response.Write "Fecha Final "		
		                CloseTd						
		                OpenTd "'labelcombo'", ""
				                Response.Write "Año:"
				                OpenCombo "end_year",  "class=listafecha"
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
				                OpenCombo "end_month",  "class=listafecha"
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
				                PlaceInput "end_day", "text", "01","id='RN             Día' class=listafecha size=2 maxlength=2"
		                CloseTd
	                CloseTr
                    opentr""
                        opentd"",""
                            Response.Write "<br><br>"
                        closetd
                    closetr		
	                OpenTr "''"
		                OpenTd "''", "align=center colspan=4"						
			                PlaceInput "Enviar","submit","Enviar","class=button-OLD "
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