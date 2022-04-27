<%@LANGUAGE="VBSCRIPT"%> 
<% 
response.buffer = true 

' Connessione al DB e variabile obj Connection ----> Cn  
dim Cn		
dim cmdTemp
		

' ------------------ connect database -----------------------
	Set Cn = Server.CreateObject("ADODB.Connection")
	Cn.ConnectionTimeout = 60
	Cn.CommandTimeout = 60
	Cn.Open "DSN=ApuliaSoft;" ',Session("RuntimeUserName"), Session("RuntimePassword")
	Set cmdTemp = Server.CreateObject("ADODB.Command")
'------------------------------------------------------------



' prendi il primo campo del primo record di una richiesta sql
Function GetDataFromDB(SQL)
	dim rs
	'response.write "<br><br> DEBUG :: " & sql
	'Response.end
	Set rs = Server.CreateObject("ADODB.Recordset")
	set rs = cn.execute(SQL)
	if rs.eof then
		GetDataFromDB = ""
	else
		GetDataFromDB = rs(0) & ""
	end if
					
	rs.close
	set rs = nothing
end Function


function ShowResult(lcStringAggregatorFields)
dim sql, rs, lcTableProduct

dim flagAggrega
dim ChrSep
	
	ChrSep = ","
	lcSqlFieldAggregator = ""
		
	lcTableProduct = "<table id=""datatable"" class=""table table-striped table-bordered"">"
	
	'Creazione testata tabella
	lcTableProduct = lcTableProduct & "<thead>"
		lcTableProduct = lcTableProduct & "<tr>"
	
		if lcStringAggregatorFields <> "" then
			flagAggrega = true
			'Aggregatori selezionati
			arrayField = split(lcStringAggregatorFields, ChrSep)
			for i = lbound(arrayField) to ubound(arrayField)
				if arrayField(i) <> "" then
					sql = "SELECT FlagAggregator FROM SY_TableFields WHERE Alias = " & arrayField(i) & " "
					lcFlagAggregator = getdatafromdb(sql) 
					if lcFlagAggregator= "S" then
						sql = "SELECT FieldNameDB FROM SY_TableFields WHERE Alias = " & arrayField(i) & " "
						lcSqlFieldAggregator = lcSqlFieldAggregator & getdatafromdb(sql) & ","
					end if
					sql = "SELECT DescrNameTB FROM SY_TableFields WHERE Alias = " & arrayField(i) & " "
					lcTableProduct = lcTableProduct & "<th>" & getdatafromdb(sql) & "</th>"
				end if
			next
			sql = "SELECT DescrNameTB FROM SY_TableFields WHERE FlagAggregator = 'N' "
			lcTableProduct = lcTableProduct & "<th>" & getdatafromdb(sql) & "</th>"
		
		else
			flagAggrega = false
			sqlSysTable = "SELECT FieldNameDB, DescrNameTB, FlagAggregator FROM SY_TableFields ORDER by Priority "
			Set rs = Server.CreateObject("ADODB.Recordset")
			cmdTemp.CommandText = sqlSysTable
			Set cmdTemp.ActiveConnection = Cn
			rs.Open cmdTemp, , 1, 1
			do until rs.eof
				if rs("FlagAggregator") = "S" then
					lcSqlFieldAggregator = lcSqlFieldAggregator & rs("FieldNameDB") & ","
				end if
				lcTableProduct = lcTableProduct & "<th>" & rs("DescrNameTB") & "</th>"
				rs.movenext
			loop

			rs.close
			set rs = nothing
		end if
		lcTableProduct = lcTableProduct & "</tr>"
	lcTableProduct = lcTableProduct & "</thead>"
	

	'Creazione lista campi aggregatori
	lcSqlFieldAggregator = left(lcSqlFieldAggregator, len(lcSqlFieldAggregator)-1)

	if flagAggrega = true then
		sql = "SELECT " & lcSqlFieldAggregator & ", SUM(TotHours) as TotHours FROM View_EmpProjectActivitiesDet "
		sql = sql & "GROUP BY " & lcSqlFieldAggregator & " "
		sql = sql & "ORDER BY " & lcSqlFieldAggregator & " "
	else
		sql = sql & "SELECT " & lcSqlFieldAggregator & ", TotHours FROM View_EmpProjectActivitiesDet ORDER BY View_EmpProjectActivitiesDet.Act_ID "
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	cmdTemp.CommandText = sql
	Set cmdTemp.ActiveConnection = Cn
	rs.Open cmdTemp, , 1, 1
	
	
		arrayField = split(lcSqlFieldAggregator, ChrSep)
		
		lcTableProduct = lcTableProduct & "<tbody>"
		do until rs.eof
			lcTableProduct = lcTableProduct & "<tr>"
				
				for i = lbound(arrayField) to ubound(arrayField)
					lcTableProduct = lcTableProduct & "<td>" & rs(i) & "</td>"
				next
				lcTableProduct = lcTableProduct & "<td>" & rs("TotHours") & "</td>"
			lcTableProduct = lcTableProduct & "</tr>"
			
			rs.movenext
		loop
		lcTableProduct = lcTableProduct & "</tbody>"
		
	rs.close
	set rs = nothing
	
	
	lcTableProduct = lcTableProduct & "</table>"	
	
	ShowResult = lcTableProduct


                      

end function



function CreateInputSelectIn(lcValueField)
dim appoHtml, rsAppo
dim sql
	
	sql = "SELECT DescrNameTB, Alias FROM SY_TableFields WHERE FlagAggregator = 'S' "
	if lcValueField <> "" then
		sql = sql & "AND Alias not in (" & lcValueField & ") "
	end if
	sql = sql & "ORDER BY Priority "
	
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	cmdTemp.CommandText = sql
	Set cmdTemp.ActiveConnection = Cn
	rs.Open cmdTemp, , 1, 1

	appoHtml = "<select name=""fieldValueIn"" class=""select2_multiple form-control"" multiple=""multiple"">"
	
	do until rs.eof
		appoHtml = appoHtml & "<option value=""" & rs("Alias") & """>" & rs("DescrNameTB") & "</option>"
	
		rs.movenext
	loop
	
	appoHtml = appoHtml & "</select>"
	
	rs.close
	set rs = nothing
	
	CreateInputSelectIn = appoHtml
                            
end function




function CreateInputSelectOut(lcValueField)
dim appoHtml, arrayValue
dim sql
	
	appoHtml = "<select name=""fieldValueOut"" class=""select2_multiple form-control"" multiple=""multiple"">"
	if lcValueField <> "" then
		lcValueField = trim(lcValueField)
		lcValueField = replace(lcValueField, "'", "")
		arrayValue = split(lcValueField, ",")
		for i = lbound(arrayValue) to ubound(arrayValue)
			'Verifica esistenza
			sql = "SELECT DescrNameTB FROM SY_TableFields WHERE Alias = '" & arrayValue(i) & "' AND FlagAggregator = 'S' "
			
			appoHtml = appoHtml & "<option value=""" & arrayValue(i)  & """>" & getdatafromdb(sql) & "</option>"
		next
	
	end if
	appoHtml = appoHtml & "</select>"
	CreateInputSelectOut = appoHtml
                  
end function


if request.form("OP") <> "" then
	select case request.form("OP")
		case "SpostaIn" 'Caso di inserimento di nuovi campi per raggruppare dati
			
			'Verifico se è un campo su cui è possibile raggruppare
			'Se corretto lo inserisco nella variabile Session
			if request("fieldValueIn") <> "" then
				arrayStringValue = request("fieldValueIn")
				arrayStringValue = trim(arrayStringValue)
				arrayStringValue = replace(arrayStringValue, " ", "")
				arrayStringValue = split(arrayStringValue, ",")
				for i = lbound(arrayStringValue) to ubound(arrayStringValue)
					sql = "SELECT Alias FROM SY_TableFields WHERE Alias = '" & arrayStringValue(i) & "' AND FlagAggregator = 'S' " 
					appoValue = getdatafromdb(sql)
					if appoValue <> "" then
						if Session("ListGroupField") <> "" then
							if instr(1, Session("ListGroupField"), arrayStringValue(i)) > 0 then
							
							else
								Session("ListGroupField") = Session("ListGroupField") & ",'" & appoValue & "'"
							end if
						else
							Session("ListGroupField") = "'" & appoValue & "'"
						end if
					end if
				next
			end if
		
		case "SpostaOut" 'Caso di eliminazione di campi dalla stringa dei campi per raggruppare
			if Session("ListGroupField") <> "" then
				if request("fieldValueOut") <> "" then
					arrayStringValue = request("fieldValueOut")
					arrayStringValue = trim(arrayStringValue)
					arrayStringValue = replace(arrayStringValue, " ", "")
					arrayStringValue = split(arrayStringValue, ",")
					for i = lbound(arrayStringValue) to ubound(arrayStringValue)
						if instr(1, Session("ListGroupField"), arrayStringValue(i)) >0 then
							Session("ListGroupField") = replace(Session("ListGroupField"), "'" & arrayStringValue(i) & "'", "")
						end if
					next
					Session("ListGroupField") = replace(Session("ListGroupField"), ",,", ",")
					if right(Session("ListGroupField"), 1) = "," then
						Session("ListGroupField") = left(Session("ListGroupField"), len(Session("ListGroupField")) - 1)
					end if
					if left(Session("ListGroupField"), 1) = "," then
						Session("ListGroupField") = right(Session("ListGroupField"), len(Session("ListGroupField")) - 1)
					end if
				end if
			end if
	
	end select
else
	Session("ListGroupField") = ""
end if

lcTableToShow = ShowResult(Session("ListGroupField"))
%>



<!DOCTYPE html>
<html lang="en">
<head>
	<!-- #include file="include_headerJscript.asp"-->
<script LANGUAGE="javascript">
function js_validate(list,field_value, flag_ctrl_field)
{
	list.OP.value = field_value;
	list.submit();
	
}
</script>	
</head>
<body class="nav-md">
<div class="container body">
	<div class="main_container">
	<br /><br />

		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="x_panel">
				<div class="x_title">
					<h2>Campi e raggruppamenti <small>Seleziona/Deseleziona i campi su cui effettuare raggruppamenti</small></h2>
					<ul class="nav navbar-right panel_toolbox">
						<li><a class="collapse-link"><i class="fa fa-chevron-up"></i></a></li>
						<li><a class="close-link"><i class="fa fa-close"></i></a></li>
					</ul>
					<div class="clearfix"></div>
				</div>
				<div class="x_content">
				<form name="functions" method="post" action="apuliasoft.asp" onSubmit="return false;">
				<input type="hidden" name="OP" value="">
				<div class="form-group">
					   <div class="col-md-3 col-sm-3 col-xs-3">
                         <%= CreateInputSelectIn(Session("ListGroupField")) %>
                        </div>
						<div class="col-md-1 col-sm-1 col-xs-1">
						<button type="submit" class="btn btn-primary" onclick="javascript:js_validate(this.form,'SpostaIn')">--></button><br><br>
                        <button type="submit" class="btn btn-success" onclick="javascript:js_validate(this.form,'SpostaOut')"><--</button>
                        </div>
						
						<div class="col-md-3 col-sm-3 col-xs-3">
                          <%= CreateInputSelectOut(Session("ListGroupField"))%>
						  
                        </div>
						 <div class="col-md-2 col-sm-2 col-xs-2"></div>
                 </div>
				 </form>
				</div>
            </div>
		</div> 
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="x_panel">
				<div class="x_title">
					<h2>Elenco attività raggruppate <small>Elenco ottenuto secondo le selezione dei campi</small></h2>
					<ul class="nav navbar-right panel_toolbox">
						<li><a class="collapse-link"><i class="fa fa-chevron-up"></i></a></li>
						<li><a class="close-link"><i class="fa fa-close"></i></a></li>
					</ul>
					<div class="clearfix"></div>
				</div>
				<div class="x_content">
					<%= lcTableToShow %>
				</div>
            </div>
		</div>
	</div>
</div>
<!-- #include file="include_FooterJScript.asp"-->
</body>
</html>
