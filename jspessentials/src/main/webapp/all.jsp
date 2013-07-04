<%@page import="java.net.URL"%>
<%@page import="org.zkoss.zk.ui.Execution"%>
<%@page import="org.zkoss.zss.api.*"%>
<%@page import="org.zkoss.zss.api.model.*"%>
<%@page import="org.zkoss.zssex.ui.ExecutionBridge"%>
<%@page import="org.zkoss.zssjsp.JSPBookProvider"%>
<%@page import="org.zkoss.zss.jspessentials.DateUtil"%>
<%@page import="org.zkoss.zss.jspessentials.DemoBookProvider"%>
<%@page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@taglib prefix="zss" uri="http://www.zkoss.org/jsp/zss"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<title>Application for Leave</title>
		<zss:zkhead/>
	</head>
<%
	//prevent page cache in browser side
	response.setHeader("Pragma", "no-cache"); 
	response.setHeader("Cache-Control", "no-store, no-cache"); 

	/*
	ExecutionBridge eb;
	new ExecutionBridge(getServletContext(),request,response){
		protected void process(Execution execution) throws Exception{
			String dateFormat =  "yyyy/MM/dd";
			URL bookUrl = _servletContext.getResource("/WEB-INF/books/application_for_leave.xlsx");
			Book book = Importers.getImporter().imports(bookUrl, "app4leave");
			
			Sheet sheet = book.getSheetAt(0);
			
			//reset sample data
			//you can use a cell reference to get a range
			Range from = Ranges.range(sheet,"E5");//Ranges.range(sheet,"From");
			//or you can use a name to get a range (the named rnage has to be set in book);
			Range to = Ranges.rangeByName(sheet,"To");
			Range reason = Ranges.rangeByName(sheet,"Reason");
			Range applicant = Ranges.rangeByName(sheet,"Applicant");
			Range requestDate = Ranges.rangeByName(sheet,"RequestDate");
			
			//use range api to set the cell data
			from.setCellEditText(DateUtil.tomorrow(0,dateFormat));
			to.setCellEditText(DateUtil.tomorrow(0,dateFormat));
			reason.setCellEditText("");
			applicant.setCellEditText("");
			requestDate.setCellEditText(DateUtil.today(dateFormat));
			
			_request.setAttribute("book",book);	
		}
	}.process();
	*/
	/*
	JSPBookProvider provider = new DemoBookProvider();
	*/
%>
<body>
	<button id="resetBtn">Reset</button>
	<button id="checkBtn">OK</button>
	<div>
		<!-- 
		2 way to initial a book, by book attribute or booProvider attribute
		book="${book}" or
		bookProvider="< %=provider%>" or
		bookProvider="org.zkoss.zss.jspessentials.DemoBookProvider"
		 -->
		<zss:spreadsheet id="myzss" 
			bookProvider="org.zkoss.zss.jspessentials.DemoBookProvider"
			width="800px" height="600px" 
			maxrows="100" maxcolumns="20"
			showToolbar="true" showFormulabar="true" showContextMenu="true" showSheetbar="true"/>
	</div>
	<script type="text/javascript">
	//jq is jquery name in zk, which version is 1.6.4 in sparedsheet 3.0.0 (zk 6.5.3 and later) 
	jq(document).ready(function(){
		//register client event on button by jquery api 
		jq("#checkBtn").click(function(){
			postAjax("check");
		});
		jq("#resetBtn").click(function(){
			postAjax("reset");
		});
	});
	
	function postAjax(action){
		//use zk client side api to get client side sparedsheet and desktop object 
		var zss = zk.Widget.$("$myzss");
		var desktop = zss.desktop;
		
		//use jquery api to post ajax to your servlet (in this demo, it is AjaxBookServlet),
		//provide desktop id and spreadsheet uuid to access zk component data in your servlet 
		jq.ajax({url:"ajaxBook",//the servlet url to handle ajax request 
			data:{dtid:desktop.id,ssuuid:zss.uuid,action:action},
			dataType:'json'}).done(handleAjaxResult);
	}
	
	//the method to handle ajax result from your servlet 
	function handleAjaxResult(result){
		//eval any zkjs for updating zk ui 
		if(result.zkjs){
			eval(result.zkjs);
		}
		
		//use your way to hanlde you ajax message or error 
		if(result.message){
			alert(result.message);
		};
		if(result.error){
			alert(result.error);
		};
		
		//use your way handle your ajax action result 
		if(result.action == "check" && result.valid){
			if(result.form){
				//create a form dynamically to submit the form data 
				var field,form = jq("<form action='submitted.jsp' method='post'/>").appendTo('body');
				for(var nm in result.form){
					field = jq("<input type='hidden' name='"+nm+"' />").appendTo(form);
					field.val(result.form[nm]);
				}
				form.submit();
			}
		};
	}
	</script>	
</body>
</html>
