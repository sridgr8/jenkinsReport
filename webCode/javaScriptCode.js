var excelFileName="Hello World.xlsx";
var excel = new ActiveXObject("Excel.Application");
var fileLoc=window.location.pathname.split('/');
var finalFileLoc="";
for(var i=1;i<fileLoc.length - 1;i++){
	finalFileLoc+=fileLoc[i]+"\\";
}
//var excel_file = excel.Workbooks.Open("C:\\Users\\srinivasulu.kummitha\\Desktop\\Scripts\\Hello World.xlsx");
var excel_file = excel.Workbooks.Open(finalFileLoc+"excelFiles\\"+excelFileName);
var excel_sheet = excel.Worksheets("Sheet1");

var xlUp = -4162;
var rowCount = (excel_sheet.cells(excel_sheet.rows.count,1).end(xlUp).row);

var passCount=0,failCount=0,otherCount=0;

var testCaseNameColNum=1, statusColNum=2, execPersonColNum=5, timeColNum=3, dateColNum=4;

var ind, dateFilterChoice=0;
var startDate, res, endDate, excelDate="";

var passCountArr=new Array();
var failCountArr=new Array();
var otherCountArr=new Array();
var myTableHeader="<table class=\"table table-striped\"><tr><td align=\"center\"><b>Test Case Name<b></td><td align=\"center\"><b>Execution Status</b></td><td align=\"center\"><b>Executed By</b></td></tr>";

var execPersonArr=new Array();
var execPersonArrCount=new Array();
var execNameListVar="<select class=\"form-control\" id=\"execNameList\" onchange=\"drawExecStatusPieChart(false,false,false); drawPersonExecCountPieChart(); radioBtnStatusReset();\"><option value=\"\Everyone\" selected=\"selected\">Everyone</option>";

function onLoadFunction() {	
  dateFilterChoice=0;
  dateFilterReset();
  
	for(var i=2;i<=rowCount;i++){
		if((excel_sheet.Cells(i,statusColNum).Value)=="Passed"){
			passCount+=1;
		}
		else if((excel_sheet.Cells(i,statusColNum).Value)=="Failed"){
			failCount+=1;
		}
		else{
			otherCount+=1;
		}
	}
	//drawExecStatusPieChart(false,false,false);
	
	for(var j=2;j<=rowCount;j++){
		if(execPersonArr.indexOf(excel_sheet.Cells(j,execPersonColNum).Value)>=0){	
			execPersonArrCount[execPersonArr.indexOf(excel_sheet.Cells(j,execPersonColNum).Value)]+=1;
		}
		else{		
			execPersonArr.push(excel_sheet.Cells(j,execPersonColNum).Value);
			execPersonArrCount[execPersonArr.indexOf(excel_sheet.Cells(j,execPersonColNum).Value)]=1;
			passCountArr.push(0);
			failCountArr.push(0);
			otherCountArr.push(0);
		}
	}  
	
	for(k=0;k<execPersonArr.length;k++){		
		execNameListVar+="<option value = \""+execPersonArr[k]+"\">"+execPersonArr[k]+"</option>";				
	}
	execNameListVar+="</select>";
	document.getElementById('execNameDiv').innerHTML = execNameListVar;
	
	for(var l=2;l<=rowCount;l++){
		ind=(execPersonArr.indexOf(excel_sheet.Cells(l,execPersonColNum).Value));
		if((excel_sheet.Cells(l,statusColNum).Value)=="Passed"){
			passCountArr[ind]+=1;
		}
		else if((excel_sheet.Cells(l,statusColNum).Value)=="Failed"){
			failCountArr[ind]+=1;
		}
		else{
			otherCountArr[ind]+=1;
		}
	}
	drawExecStatusPieChart(false,false,false);
	drawPersonExecCountPieChart();
}

function dateFilterReset(){
	var options = {month: '2-digit',day: '2-digit',year: 'numeric'};
	let today = new Date().toLocaleDateString('en-US',options);
	document.querySelector("#dateStart").value = today;
	document.querySelector("#dateEnd").value = today;
}

function showAllPlans(){
	
	var count=0;
	var myTable= myTableHeader;
	var e = document.getElementById("execNameList");
	if(dateFilterChoice==0){
		for(var i=2;i<=rowCount;i++){
			if((e.options[e.selectedIndex].value)=="Everyone"){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else{
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}
	}
	else if(dateFilterChoice==1){	
		for(var i=2;i<=rowCount;i++){
			if((e.options[e.selectedIndex].value)=="Everyone" && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else{
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value) && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}	
	}
	else if(dateFilterChoice==2){
		for(var i=2;i<=rowCount;i++){
			if((e.options[e.selectedIndex].value)=="Everyone" && (excelDate >= startDate) && (excelDate <= endDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else{
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value) && (excelDate >= startDate) && (excelDate <= endDate)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}
	}
	myTable+= "</table>";
	if(dateFilterChoice==3){
		document.getElementById('contentTable').innerHTML = "<center><b>Please Select a Date Filter</b></center>";
	}
	else{
	if(count!=0){
		document.getElementById('contentTable').innerHTML = myTable;
	}
	else{
		document.getElementById('contentTable').innerHTML = "<center><b>No Content Found for the Selected Filters</b></center>";
	}
	}
	drawExecStatusPieChart(false,false,false);
}

function showPassedPlans(){
	
	var status,count=0;
	var myTable= myTableHeader;
	var e = document.getElementById("execNameList");
	if(dateFilterChoice==0){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status=="Passed"){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status=="Passed"){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}
	}
	else if(dateFilterChoice==1){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status=="Passed" && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status=="Passed" && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}
	}
	else if(dateFilterChoice==2){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status=="Passed" && (excelDate >= startDate) && (excelDate <= endDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status=="Passed" && (excelDate >= startDate) && (excelDate <= endDate)){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}
	}
	myTable+= "</table>";
	if(dateFilterChoice==3){
		document.getElementById('contentTable').innerHTML = "<center><b>Please Select a Date Filter</b></center>";
	}
	else{
	if(count!=0){
		document.getElementById('contentTable').innerHTML = myTable;
	}
	else{
		document.getElementById('contentTable').innerHTML = "<center><b>No Content Found for the Selected Filters</b></center>";
	}
	}
	drawExecStatusPieChart(true,false,false);
}

function showFailedPlans(){
	
	var status,count=0;
	var myTable= myTableHeader;
	var e = document.getElementById("execNameList");
	if(dateFilterChoice==0){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status=="Failed"){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status=="Failed"){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}
	}
	else if(dateFilterChoice==1){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status=="Failed" && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status=="Failed" && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
	}
	}
	else if(dateFilterChoice==2){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status=="Failed" && (excelDate >= startDate) && (excelDate <= endDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status=="Failed" && (excelDate >= startDate) && (excelDate <= endDate)){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
	}
	}
	myTable+= "</table>";
	if(dateFilterChoice==3){
		document.getElementById('contentTable').innerHTML = "<center><b>Please Select a Date Filter</b></center>";
	}
	else{
	if(count!=0){
		document.getElementById('contentTable').innerHTML = myTable;
	}
	else{
		document.getElementById('contentTable').innerHTML = "<center><b>No Content Found for the Selected Filters</b></center>";
	}
	}
	drawExecStatusPieChart(false,true,false);
}

function showOtherPlans(){
		
	var status,count=0;
	var myTable= myTableHeader;
	var e = document.getElementById("execNameList");
	if(dateFilterChoice==0){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status!="Passed" && status!="Failed"){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status!="Passed" && status!="Failed"){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
		}
	}
	else if(dateFilterChoice==1){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status!="Passed" && status!="Failed" && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status!="Passed" && status!="Failed" && ((excel_sheet.Cells(i,dateColNum).Value) == startDate)){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
	}
	}
	else if(dateFilterChoice==2){
		for(var i=1;i<=rowCount;i++){
			status=(excel_sheet.Cells(i,statusColNum).Value);
			if(((e.options[e.selectedIndex].value)=="Everyone") && status!="Passed" && status!="Failed" && (excelDate >= startDate) && (excelDate <= endDate)){
				myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
				count=1;
			}
			else if(((e.options[e.selectedIndex].value)!="Everyone") && status!="Passed" && status!="Failed" && (excelDate >= startDate) && (excelDate <= endDate)){
				if((e.options[e.selectedIndex].value)==(excel_sheet.Cells(i,execPersonColNum).Value)){
					myTable+="<tr><td align=\"left\">"+(excel_sheet.Cells(i,testCaseNameColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,statusColNum).Value)+"</td><td align=\"center\">"+(excel_sheet.Cells(i,execPersonColNum).Value)+"</td></tr>";
					count=1;
				}
			}
	}
	}
	myTable+= "</table>";
	if(dateFilterChoice==3){
		document.getElementById('contentTable').innerHTML = "<center><b>Please Select a Date Filter</b></center>";
	}
	else{
	if(count!=0){
		document.getElementById('contentTable').innerHTML = myTable;
	}
	else{
		document.getElementById('contentTable').innerHTML = "<center><b>No Content Found for the Selected Filters</b></center>";
	}
	}
	drawExecStatusPieChart(false,false,true);
}

function drawExecStatusPieChart(passedExplode,failedExplode,otherExplode){
	document.getElementById('statusPieChartContainer').innerHTML = "";
	var e = document.getElementById("execNameList");
	var selectedUser=(e.options[e.selectedIndex].value);
	
	if(selectedUser=="Everyone"){
		var data = [
					{x: "Passed", value: passCount, exploded: passedExplode, fill:"#5cb85c"},
					{x: "Failed", value: failCount, exploded: failedExplode, fill:"#d9534f"},
					{x: "Others", value: otherCount, exploded: otherExplode, fill:"#ec971f"}
		];
	}
	else{
		data = [
					{x: "Passed", value: passCountArr[execPersonArr.indexOf(selectedUser)], exploded: passedExplode, fill:"#5cb85c"},
					{x: "Failed", value: failCountArr[execPersonArr.indexOf(selectedUser)], exploded: failedExplode, fill:"#d9534f"},
					{x: "Others", value: otherCountArr[execPersonArr.indexOf(selectedUser)], exploded: otherExplode, fill:"#ec971f"}
		];
	}
	
	var chart = anychart.pie();
	chart.title("Test Plan Execution Status");
	
	chart.data(data);
	
	// display the chart in the container
	chart.container('statusPieChartContainer');
	//chart.fill("aquastyle");
	//art.labels().position("outside");
	//art.connectorStroke({color: "#595959", thickness: 2, dash:"2 2"});
	chart.draw();
	// set legend position
	chart.legend().position("right");
	// set items layout
	chart.legend().itemsLayout("vertical");
	// sort elements
	chart.sort("desc");
}

function drawPersonExecCountPieChart(){
	document.getElementById('numExecPieChartContainer').innerHTML = "";
	var e = document.getElementById("execNameList");
	var selectedUser=(e.options[e.selectedIndex].value);
	var data=new Array();
	for (var i=0;i<execPersonArr.length;i++){
			data[i] = {x: execPersonArr[i], value: execPersonArrCount[i], exploded: (selectedUser==execPersonArr[i])};
	}
	var chart = anychart.pie();
	chart.title("Test Plan Execution Numbers");
	chart.data(data);
	// display the chart in the container
	chart.container('numExecPieChartContainer');
	//chart.fill("aquastyle");
	chart.labels().position("outside");
	chart.connectorStroke({color: "#595959", thickness: 1, dash:"2 2"});
	chart.draw();
	// set legend position
	chart.legend().position("right");
	// set items layout
	chart.legend().itemsLayout("vertical");
	// sort elements
	chart.sort("desc");
}

function selectDate(){
	dateFilterChoice=1;
	startDate=document.getElementById('dateStart').value;
	var res = startDate.split("/");
	startDate=res[1]+"-"+res[0]+"-"+res[2];
}

function selectDateRange(){
	dateFilterChoice=2;
	startDate=new Date(document.getElementById('dateStart').value);
	endDate=new Date(document.getElementById('dateEnd').value);
	for(var i=2;i<=rowCount;i++){
		excelDate=(excel_sheet.Cells(i,dateColNum).Value);
		res=excelDate.split("-");
		excelDate=new Date(res[1]+"/"+res[0]+"/"+res[2])		
	}
}

function allFilterReset(filterValue){
	dateFilterChoice=0;
	dateFilterReset();
	document.getElementById('execNameList').selectedIndex = 0;
	showAllPlans();
	drawPersonExecCountPieChart();
	radioBtnReset(filterValue);
}
function radioBtnReset(filterValue){
	dateFilterChoice=filterValue;
	document.getElementById("lbl5").classList.remove("active");
	document.getElementById("lbl6").classList.remove("active");
	document.getElementById("lbl7").classList.remove("active");
	radioBtnStatusReset();
}

function radioBtnStatusReset(){
	document.getElementById("lbl1").classList.remove("active");
	document.getElementById("lbl2").classList.remove("active");
	document.getElementById("lbl3").classList.remove("active");
	document.getElementById("lbl4").classList.remove("active");
	document.getElementById('contentTable').innerHTML = "";
}
