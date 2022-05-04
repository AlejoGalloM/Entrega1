<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:						Sel_rep_Investor.asp 
'Path:							/operations/radication
'Created By:					jimmy Ospino Huertas
'Last Modified:
'				
'				
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
Authorize 0,15
Response.Write "<link rel='stylesheet' href='../../css/OLDMutualStyle.css' type='text/css'>" & vbCrLf
'== declares ===
Dim Inicio
Dim Fin
Dim Section
Dim Cn
Dim rs
Dim sql
Dim I
Dim J
Dim Sel
Dim arrProduct
Dim arrUser
Dim ProductCombo
Dim UserCombo


'== initials asignments ==
set cn = getconnpipelinedb
set rs = Server.CreateObject("ADODB.RecordSet") 
ProductCombo = PlaceDocTypeCombo ("bttntext", cn, "Todos")

sql = "parameters..Product_GetList" 
rs.Open sql,cn,3
If rs.BOF And rs.EOF Then
	arrProduct = 0
Else
	arrProduct = rs.GetRows()
End If
rs.Close 

write_sp_log cn, 11200, "parameters..Product_GetList", 0, "", "", 0, 0, "", "sel_rep_Investor.asp " & _
"Loaded by " & Session("sp_miLogin")

set cn = getconnpipelinedb
set rs = Server.CreateObject("ADODB.RecordSet") 
UserCombo = PlaceDocTypeCombo ("bttntext", cn, "Todos")

sql = "sigscg..UserbyRol_GetList" 
rs.Open sql,cn,3
If rs.BOF And rs.EOF Then
	arrUser = 0
Else
	arrUser = rs.GetRows()
End If
rs.Close 

write_sp_log cn, 11200, "Usuarios..UserbyRol_GetList", 0, "", "", 0, 0, "", "sel_rep_Investor.asp " & _
"Loaded by " & Session("sp_miLogin")

set rs = Nothing
closeconnpipelinedb
set cn= Nothing

write_dataLog Response.Status,"sel_rep_Investor.asp", "sel_rep_Investor.asp Loaded by " & Session("sp_miLogin"),Session("sp_miLogin"),"parameters..Product_GetList - sigscg..UserbyRol_GetList" ,"N/A","null","Consulta","N/A"

OpenHTML
OpenHead
%>
<SCRIPT LANGUAGE=javascript src=../_pipeline_scripts/validation.js></SCRIPT>
<SCRIPT LANGUAGE=javascript>
function Validate(theForm){
  if (parseInt(theForm.Inicio.value) >= parseInt(theForm.Fin.value)) {
   alert("Número Inicial debe ser menor al Final");
   return false;
 }
 return true;
}

function ValidateDate(theform){
 if (!dateValidation(theform.start_year.value,theform.start_month.value,theform.start_day.value)){     return false;
 }
 if (!dateValidation(theform.end_year.value,theform.end_month.value,theform.end_day.value)){
     return false;
 }
 	  
 if (!(  ( ( 0 < parseInt(theform.end_min.value) ) &&  
	     ( parseInt(theform.end_min.value) < 60 ) 
	    ) &&
	   ( ( 0 < parseInt(theform.start_min.value) ) &&  
	     ( parseInt(theform.start_min.value) < 60 ) 
	   )
	  ))
	  {
   alert("Minutos entre 0 y 60");
   return false;
 }

 if (!(  ( ( 0 < parseInt(theform.end_hour.value) ) &&  
	     ( parseInt(theform.end_hour.value) < 24 ) 
	    ) &&
	   ( ( 0 < parseInt(theform.start_hour.value) ) &&  
	     ( parseInt(theform.start_hour.value) < 24 ) 
	   )
	  )==true)
	  {
   alert("Horas entre 0 y 24");
   return false;
 }
 
 var startD;
 var endD;
 startD = new Date(theform.start_year.value + " "+ theform.start_month.value+ " "+theform.start_day.value+ " "+theform.start_hour.value+ ":"+theform.start_min.value+ ":"+"00");
 endD = new Date(theform.end_year.value+ " "+ theform.end_month.value+ " "+ theform.end_day.value+ " "+ theform.end_hour.value+ ":"+ theform.end_min.value+ ":"+ "00");
 
 if (Date.parse(endD) <= Date.parse(startD) ){
   alert("Fecha Final debe ser mayor o igual a Fecha Inicial");
   return false;
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
		<div class="subtituloPagina">
			Reporte de Clientes Nuevos y/o Adición de productos
		</div>
<%

otbl"tblcontenido"
    opentr""
        opentd"",""
            Response.Write "<br>"
        closetd
    closetr
    opentr""
        opentd"",""
            otbl"tblvalores"
                OpenForm "theForm", "post", "rep_investor.asp", "onSubmit='return ValidateDate(this)&& formValidation(this)'"			
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
				                OpenCombo "start_month",  "class=listagenerica"
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
				                PlaceInput "start_day", "text", "01","id='RN             Día' class=listagenerica size=2 maxlength=2"
				                Response.Write "Hora:"				
				                PlaceInput "start_hour", "text", "01","id='RN             Hora' class=listagenerica size=2 maxlength=2"
				                Response.Write "Minuto:"				
				                PlaceInput "start_min", "text", "01","id='RN             Minuto' class=listagenerica size=2 maxlength=2"				
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
				                OpenCombo "end_month",  "class=listagenerica"
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
				                PlaceInput "end_day", "text", "01","id='RN             Día' class=listagenerica size=2 maxlength=2"				
				                Response.Write "Hora:"				
				                PlaceInput "end_hour", "text", "01","id='RN             Hora' class=listagenerica size=2 maxlength=2"
				                Response.Write "Minuto:"				
				                PlaceInput "end_min", "text", "01","id='RN             Minuto' class=listagenerica size=2 maxlength=2"				
		                CloseTd
	                CloseTr	
		            OpenTr ""
			            OpenTd "'labels'", "align=left"
				            Response.Write "Producto "		
			            CloseTd			
			            OpenTd "'labelcombo'", "align=left colspan=4"						
				            If IsArray(arrProduct) Then	
					            OpenCombo "idProduct",  "class='listagenerica' style='width:35em;'"
					            PlaceItem "", "-1","Todos"
					            For J = 0 To UBound(arrProduct, 2)												
						            PlaceItem "", arrProduct(1,J), arrProduct(2,J)								
					            Next
					            CloseCombo
				            end if 
			            CloseTd
		            CloseTr	
		            OpenTr ""
			            OpenTd "'labels'", "align=left "
				            Response.Write "Usuario "		
			            CloseTd						
	
			            OpenTd "'labelcombo'", "align=left colspan=4"						
				            If IsArray(arrUser) Then	
					            OpenCombo "idUser",  "class='listagenerica' style='width:35em;'"
					            PlaceItem "", "-1","Todos"
					            For J = 0 To UBound(arrUser, 2)												
						            PlaceItem "", arrUser(1,J), arrUser(2,J)								
					            Next
					            CloseCombo
					         end if
			            CloseTd
		            CloseTr
                    opentr""
                        opentd"",""
                            Response.Write "<br>"
                        closetd
                    closetr	
	                OpenTr ""
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