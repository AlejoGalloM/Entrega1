<%@ Language=VBScript %>
<%
'===================================================================================
'File Name:		Param_Password_Change.asp  
'Path:				login/
'Created By:		I&T - 2010/09/06
'Last Modified:	
'Additional Information: None
'===================================================================================
Option Explicit
'On Error Resume Next
Response.Buffer = True
Response.CacheControl="no-cache"
Response.AddHeader "pragma", "no-cache"
Response.Expires= -1
%>
<!--#include file="../../operations/_pipeline_scripts/tags.asp"-->
<!--#include file="../../operations/_pipeline_scripts/pipeline_scripts.asp"-->
<%
dim sql,cn
dim rsParam
dim divVisibility
dim initial_portal 
dim rsActualHisto
dim regularExpresion
dim vigenciaMax
dim vigenciaMin
dim historicoClaves
dim regularExpRepeat
dim arrParam
dim maxLength
dim minLength
dim cantidadHisto
dim regularSpace
dim regularNumeric
dim regularNumericUpper
dim resul
dim valMaxLength, valMinLength
dim valMaxVig, valMinVig
dim valMaxHisto , valRegExp
dim valRegExpRepeat, valPortal
dim strExpress, strExpressRep
dim maxAttempts
dim pwdTempExpira
Dim processInfo, component_id, valueOldLog
set initial_portal = request.item("initial_portal")

set cn = GetConnpipelineDB

if initial_portal <> 0 then
	divVisibility = "block"
else
	divVisibility ="none"
end if

if divVisibility="block" then
'Consultar los parametros

	' set rsParam = Server.CreateObject("ADODB.RecordSet") 
	' sql = "usuarios..Get_ParametroClave " & initial_portal
	' rsParam.Open sql,cn

	Set pcmd = Server.CreateObject("ADODB.Command")
	pcmd.CommandText = "usuarios..Get_ParametroClave"
	pcmd.CommandType = 4 'adCmdStoredProc
	pcmd.ActiveConnection = cn

	pcmd.Parameters.Append(pcmd.CreateParameter("@sitio", 3, 1, , initial_portal)) '3=adInteger, 1=adInput

	Set rsParam = pcmd.Execute

	If rsParam.BOF And rsParam.EOF Then
		arrParam = 0
	Else
		arrParam = rsParam.GetRows()
	End If
	rsParam.Close 
	If IsArray(arrParam) Then
		regularExpresion = arrParam(9,0)
		vigenciaMax = arrParam(0,0)
		vigenciaMin = arrParam(1,0)
		historicoClaves = arrParam(7,0)
		regularExpRepeat = arrParam(10,0)
		maxLength = arrParam(11,0)
		minLength = arrParam(2,0)
		regularSpace = arrParam(13,0)
		regularNumeric = arrParam(14,0)
		regularNumericUpper = arrParam(15,0)
		maxAttempts = arrParam(16,0)
		pwdTempExpira = arrParam(17,0)
	end if
End if

component_id = "Param_Password_Change.asp"
processInfo =  ""

valueOldLog = "vigenciaMax: "&vigenciaMax&", vigenciaMin: "&vigenciaMin&", maxLength: "&maxLength&", minLength: "&minLength&", maxAttempts: "&maxAttempts&", pwdTempExpira: "&pwdTempExpira

write_dataLog Response.Status,component_id,processInfo,Session.contents("idworker"), "usuarios..Get_ParametroClave " & initial_portal ,"",valueOldLog,"Operación-Modificación","N/A"


Sub OpenTextAreaExp(Name, Rows, Cols, ClassName, Params, Value)
	Response.Write "<TEXTAREA rows=" & Rows & " cols=" & Cols & " name=" & Name & " " & Params & " class=" & _
	ClassName & ">" & Value & vbCrLf
End Sub

Sub CloseTextAreaExp
	Response.Write "</TEXTAREA>" & vbCrLf
End Sub

function compareExpresion(Expresion2)

	if (Instr(regularExpresion,Expresion2)>0) then
		 resul = "checked"
	else
		 resul = ""
	end if	
	 compareExpresion = resul 
end function 

function compareExpresionRepeat(Expresion1)
	if (Expresion1=regularExpRepeat) then
		resul = "checked"
	else
		resul = ""
	end if
	 compareExpresionRepeat = resul 
end function


%>
<script language="javascript">
	function validateForm(form)
	{
		if (form.vigenciaMax.value == null || form.vigenciaMax.value == ''){
			alert("La vigencia m�xima no puede estar vacia");
			form.vigenciaMax.focus();
			return false;
		}else
		{
			if (form.vigenciaMax.value == '0' || parseInt (form.vigenciaMax.value) < 0){
			alert("La vigencia m�xima debe ser mayor a 0");
			form.vigenciaMax.focus();
			return false;
			}
		}
		if (form.vigenciaMin.value == null || form.vigenciaMin.value == ''){
			alert("La vigencia minima no puede estar vacia");
			form.vigenciaMin.focus();
			return false;
		}else
		{
			if (form.vigenciaMin.value == '0' || parseInt(form.vigenciaMin.value) < 0  ){
				alert("La vigencia minima debe ser mayor a 0");
				form.vigenciaMin.focus();
				return false;
			}
		}
		if (form.longitudMin.value == null || form.longitudMin.value == ''){
			alert("La longitud minima no puede estar vacia");
			form.longitudMin.focus();
			return false;
		}else
		{
			if (form.longitudMin.value == '0' || parseInt(form.longitudMin.value) < 0 ){
				alert("La longitud minima debe ser mayor a 0 ");
				form.longitudMin.focus();
				return false;
			}
		}
		if (form.longitudMax.value == null || form.longitudMax.value == ''){
			alert("La longitud minima no puede estar vacia");
			form.longitudMax.focus();
			return false;
		}else
		{
			if (form.longitudMin.value == '0' || parseInt(form.longitudMin.value) < 0 ){
				alert("La longitud maxima debe ser mayor a 0 ");
				form.longitudMin.focus();
				return false;
			}
		}
		
		if(form.valHisto.value == null || form.valHisto.value ==''){
			alert("El valor del historico no puede estar vacia");
			form.valHisto.focus();
			return false;
		}else{
			if (form.valHisto.value == '0' || parseInt(form.valHisto.value) < 0){
				alert("El valor del historico debe ser mayor a 0");
				form.valHisto.focus();
				return false;
			}
		}
		if(form.valAttempts.value == null || form.valAttempts.value ==''){
			alert("El valor del n�mero m�ximo de intentos no puede estar vacia");
			form.valAttempts.focus();
			return false;
		}else{
			if (form.valAttempts.value == '0' || parseInt(form.valAttempts.value) < 0){
				alert("El valor del n�mero m�ximo de intentos debe ser mayor a 0");
				form.valAttempts.focus();
				return false;
			}
		}
		if(this.validateNumber(form.valHisto.value)==false){
			alert("El valor del historico debe ser un n�mero entero");
			form.valHisto.focus();
			return false;
		}
		if(this.validateNumber(form.longitudMin.value)==false){
			alert("El valor de la longitud minima debe ser un n�mero entero");
			form.longitudMin.focus();
			return false;
		}
		if(this.validateNumber(form.longitudMax.value)==false){
			alert("El valor de la longitud maxima debe ser un n�mero entero");
			form.longitudMax.focus();
			return false;
		}
		if(this.validateNumber(form.vigenciaMax.value)==false){
			alert("El valor de la  vigencia maxima debe ser un n�mero entero");
			form.vigenciaMax.focus();
			return false;
		}
		if(this.validateNumber(form.vigenciaMin.value)==false){
			alert("El valor de la  vigencia minima debe ser un n�mero entero");
			form.vigenciaMin.focus();
			return false;
		}
		
		if(this.validateNumber(form.valAttempts.value)==false){
			alert("El valor del n�mero m�ximo de intentos debe ser un n�mero entero");
			form.vigenciaMin.focus();
			return false;
		}
		
		if(this.validateIsMayor(form.vigenciaMax.value,form.vigenciaMin.value)==false){
			alert("La vigencia m�xima debe ser mayor a la minima");
			form.vigenciaMax.focus();
			return false;
		}
		
		if(this.validateIsMayor(form.longitudMax.value,form.longitudMin.value)==false){
			alert("La longitud m�xima debe ser mayor a la minima");
			form.longitudMax.focus();
			return false;
		}

		form.accion.value="1"; 
		return true;		
	}
	
	
	
	function validateNumber(value)
	{
		var vAux = parseInt(value);
		if (isNaN(vAux)){
			return false;
		}else
		{
			return true;
		}
	}
	
	function validateIsMayor(value1,value2){
		var vValue1 = parseInt(value1);
		var vValue2 = parseInt(value2);
		if( vValue1 > vValue2){
			return true;
		}else{
			return false;
		}
	}
</script> 
<%
OpenHTML 'HTML
	OpenHead ' HEADER
		PlaceTitle "Change Password Param"
		PlaceLink "REL", "stylesheet", "../../css/style1.css", "text/css"
		PlaceLink "REL", "stylesheet", "../../css/style.css", "text/css"
		PlaceMeta "expires", "", "Wednesday, 27-Dec-95 05:29:10 GMT"
		PlaceMeta "Pragma", "", "no_cache"
	CloseHead
	OpenBody "loginbody", "bgColor=#ffffff leftMargin=0 topMargin=0" 'Body
	OpenForm "login", "post", "Param_Password_Change.asp", "onSubmit='return validateForm(this)'"
	PlaceInput "accion", "hidden", "0", ""
		OpenTable "60%", "'' cellpadding=0 cellspacing=0 border=0 align=center" 'Table
			OpenTr ""
				OpenTd "", "&nbsp;"
					Response.Write "<br>"		
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "", "&nbsp;"
					Response.Write "<br>"			
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "", "&nbsp;"
					Response.Write "<br>"				
				CloseTd
			CloseTr
			OpenTr "class=status"
				OpenTd "",""
					Response.Write "<h3>Parametros de Configuraci�n</h3>"
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "", "&nbsp;"
					Response.Write "<br>"				
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "", "&nbsp;"
					Response.Write "<br>"				
				CloseTd
			CloseTr
			OpenTr ""
			CloseTr
			OpenTr ""
				OpenTd "", ""
					OpenTable "60%", "'' cellpadding=0 cellspacing=0 border=0"
						OpenTr ""
							OpenTd "", "&nbsp;"
								
							CloseTd
							OpenTd "'txfields'", "width=50% align=left height=37"
								Response.Write "Seleccione el Portal :"
							CloseTd
							OpenTd "", "&nbsp;"
								OpenCombo "initial_portal"," onChange='this.form.submit();'"
									if initial_portal = 0 then
										PlaceItem " selected", "0","Seleccione ..."
									else 
										PlaceItem "", "0","Seleccione ..."
									end if
									
									if initial_portal = 1 then
										PlaceItem " selected", "1","Portal Clientes"
									else 
										PlaceItem " ", "1","Portal Clientes"
									end if
									
									if initial_portal = 2 then
										PlaceItem " selected", "2","Portal Pipeline"
									else 
										PlaceItem "", "2","Portal Pipeline"
									end if									
								CloseCombo
							CloseTd
							OpenTd "", "&nbsp;"
							CloseTd						
						CloseTr
					CloseTable
				CloseTd
			CloseTr
		CloseTable
		OpenDiv "" ,"div1" , "style=display:"&divVisibility&""
		OpenTable "60%", "'' cellpadding=0 cellspacing=0 border=0 align=center" 'Table
			OpenTr ""
				OpenTd "", ""
					OpenTable "100%", "'' cellpadding=0 cellspacing=0 border=0"
						OpenTr ""
							OpenTd "'txfields'", "width=70% align=left height=37"
								Response.Write "Duraci&oacute;n M&aacute;xima de la contrase�a en d&iacute;as"
							CloseTd
							OpenTd "'txfields'", "width=70% align=left height=37"
								PlaceInput "vigenciaMax", "text", vigenciaMax, "class=fields id='RVigenciaMax'"
							CloseTd
						CloseTr
						OpenTr ""
							OpenTd "'txfields'", "width=70% align=left height=37"
								Response.Write "Cantidad permitida de cambios de contrase�a por d�a"
							CloseTd
							OpenTd "'txfields'", "width=70% align=left height=37"
								PlaceInput "vigenciaMin", "text", vigenciaMin, "class=fields id='RVigenciaMin'"
							CloseTd
						CloseTr
						OpenTr ""
							OpenTd "'txfields'", "width=70% align=left height=37"
								Response.Write "M&aacute;ximo de caracteres permitidos en la contrase�a:"
							CloseTd
							OpenTd "'txfields'", "width=70% align=left height=37"
								PlaceInput "longitudMax", "text", maxLength, "class=fields id='RLongitudMax'"
							CloseTd
						CloseTr
						OpenTr ""
							OpenTd "'txfields'", "width=70% align=left height=37"
								Response.Write "M&iacute;nimo de caracteres permitidos en la contrase�a:"
							CloseTd
							OpenTd "'txfields'", "width=70% align=left height=37"
								PlaceInput "longitudMin", "text", minLength, "class=fields id='RLongitudMin'"
							CloseTd
						OpenTr ""
							OpenTd "'txfields'", "width=50% align=left height=37"
								Response.Write "Cantidad en el historico:"
							CloseTd
							OpenTd "'txfields'", "width=50% align=left height=37"
								PlaceInput "valHisto", "text", historicoClaves, "class=fields id='RvalHisto'"
							CloseTd
						CloseTr
						OpenTr ""
							OpenTd "'txfields'", "width=50% align=left height=37"
								Response.Write "N&uacute;mero m&aacute;ximo de intentos:"
							CloseTd
							OpenTd "'txfields'", "width=50% align=left height=37"
								PlaceInput "valAttempts", "text", maxAttempts, "class=fields id='RvalAttempts'"
							CloseTd
						CloseTr
						OpenTr ""
							OpenTd "'txfields'", "width=50% align=left height=37"
								Response.Write "Duraci�n M�xima de la contrase�a en horas:"
							CloseTd
							OpenTd "'txfields'", "width=50% align=left height=37"
								PlaceInput "valPwdExp", "text", pwdTempExpira, "class=fields id='RvalAttempts'"
							CloseTd
						CloseTr
						CloseTr
						OpenTr ""
							OpenTd "'txfields'", "width=50% align=left height=50"
								Response.Write "Validaci&oacute;n de caracteres:"
							CloseTd
							OpenTd "'txfields'", "width=100% align=left height=37"								
								PlaceInput "checkBoxNum","checkbox", "1",compareExpresion("(?=.*\d)")
								Response.write "N&uacute;meros</br>"
							
								PlaceInput "checkBoxMin","checkbox", "1",compareExpresion("(?=.*[a-z])")
								Response.write "Min&uacute;sculas</br>"
						
								PlaceInput "checkBoxMay","checkbox", "1",compareExpresion("(?=.*[A-Z])")
								Response.write "May&uacute;sculas</br>"

								PlaceInput "checkBoxEsp","checkbox", "1",compareExpresion("(?=.*[!@#$%^&*()_+}{"":;?/>.<,])")
								Response.write "C. especiales</br>"
							CloseTd
						
						CloseTr
						OpenTr ""
							OpenTd "'txfields'", "width=70% align=left height=37"
								Response.Write "Verificar Caracteres repetidos:"
							CloseTd
							OpenTd "'txfields'", "width=50% align=left height=37"
								PlaceInput "checkBoxRep","checkbox", "1",compareExpresionRepeat("(.)\1{1,}")
								Response.write "Si</br>"
								
							CloseTd
						CloseTr						
					CloseTable
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "", "&nbsp;"
					Response.Write "<br>"				
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "", "&nbsp;"
					Response.Write "<br>"				
				CloseTd
			CloseTr
			OpenTr ""
				OpenTd "''", "width=50% align=right"
					PlaceInput "btnSend", "submit", "Enviar", "id='Enviar' class=sbttn2"				
				CloseTd
			CloseTr		
		CloseTable
		CloseDiv
		OpenTable "60%", "'' cellpadding=0 cellspacing=0 border=0 align=center" 'Table
			OpenTr ""
				if   Request.Item("ok")<>"" or Request.Item("ok")="1"  then
								OpenTd "'txfields'", "width=80% align=left height=37"
									Response.Write "Configuraci�n actualizada con �xito"
								CloseTd
				end if
			CloseTr
		CloseTable
	CloseForm
	CloseBody	'Close body
CloseHTML

valMaxLength = Request.Item("longitudMax")
valMinLength = Request.Item("longitudMin")

strExpress = "(?=^.{" & valMinLength & "," & valMaxLength & "}$)"

if Request.Item("checkBoxNum") <> "" then
	strExpress = strExpress & "(?=.*\d)"
end if
if Request.Item("checkBoxMin") <> "" then
	strExpress = strExpress & "(?=.*[a-z])"
end if
if Request.Item("checkBoxMay") <> "" then
	strExpress = strExpress & "(?=.*[A-Z])"
end if
if Request.Item("checkBoxEsp") <> "" then
	strExpress = strExpress & "(?=.*[!@#$%^&*()_+}{"":;?/>.<,])"
end if
strExpress = strExpress & "(?!.*\s).*$"


if Request.Item("checkBoxRep") <> "" then
	strExpressRep = "(.)\1{1,}"
else
	strExpressRep = "(.)\1{"&valMaxLength&",}"
end if

if  Request.Item("accion") = "1" then
	'Se procede a actualizar el registro
	valMaxLength = Request.Item("longitudMax")
	valMinLength = Request.Item("longitudMin")
	valMaxVig = Request.Item("vigenciaMax")
	valMinVig = Request.Item("vigenciaMin")
	valMaxHisto = Request.Item("valHisto")
	valRegExp = Request.Item("valExpreRegular")
	valRegExpRepeat = Request.Item("valExpreRegRep")
	valPortal = Request.Item("initial_portal")
	MaxAttempts = Request.Item("valAttempts")
	pwdTempExpira = Request.Item("valPwdExp")
	sql = "EXEC usuarios..Upd_SecurityParams '" & valMaxVig &"','" & valMinVig &"','" & valMinLength& "','" & valMaxHisto &"'"
	sql = sql & ",'" & strExpress &"','"& strExpressRep & "','" & valMaxLength & "','" & MaxAttempts& "','" & pwdTempExpira & "','"  & valPortal & "'"
	cn.Execute sql
	If Err.number <> 0 Then
		Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
		"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
		Server.URLEncode(Err.source))
	end if
	CloseConnpipelineDB
	response.redirect "Param_Password_Change.asp?ok=1"
End If

CloseConnpipelineDB
If Err.number <> 0 Then
	Response.Redirect(Application("ErrorURL") & "?page=" & Server.URLEncode(Request.ServerVariables("URL")) & _
	"&ErrNum=" & Err.number & "&ErrDesc=" & Server.URLEncode(Err.description) & "&ErrSource=" & _
	Server.URLEncode(Err.source))
End If
%>