<%@page import="java.security.Principal"%>
<%@page import="org.zkoss.zss.jspessentials.DateUtil"%>
<%@page import="org.zkoss.zss.api.Ranges"%>
<%@page import="org.zkoss.zss.api.model.Sheet"%>
<%@page import="org.zkoss.zk.ui.Execution"%>
<%@page import="org.zkoss.zssex.ui.ExecutionBridge"%>
<%@page import="java.net.URL"%>
<%@page import="org.zkoss.zss.api.Importers"%>
<%@page import="org.zkoss.zss.api.model.Book"%>
<%@page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@taglib prefix="zss" uri="http://www.zkoss.org/jsp/zss"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<zss:zkhead/>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>My First ZK Spreadsheet JSP application</title>
</head>
<%
	ExecutionBridge eb;
	new ExecutionBridge(getServletContext(),request,response){
		protected void process(Execution execution) throws Exception{
			URL bookUrl = _servletContext.getResource("/WEB-INF/books/application_for_leave.xlsx");
			Book book = Importers.getImporter().imports(bookUrl, "mybook");
			
			Sheet sheet = book.getSheetAt(0);
			
			//reset sample data
			Ranges.range(sheet,"D4").setCellEditText(DateUtil.tomorrow(0,"yyyy/MM/dd"));//from
			Ranges.range(sheet,"D5").setCellEditText(DateUtil.tomorrow(0,"yyyy/MM/dd"));//to
			Ranges.range(sheet,"C6").setCellEditText("");//reason
			
			Principal p = _request.getUserPrincipal();
			if(p!=null){
				Ranges.range(sheet,"C7").setCellEditText(p.getName());//applicant name for base servlet auth
			}else{
				Ranges.range(sheet,"C7").setCellEditText("");//applicant name for base servlet auth
			}
			
			_request.setAttribute("book",book);	
		}
	}.process();
%>
<body>
	<button id="resetBtn">Reset</button>
	<button id="checkBtn">Check</button>
	
	<div width="100%" height="100%">
		<zss:spreadsheet id="myzss" book="${book}"
			width="800px" height="600px" 
			maxrows="100" maxcolumns="20"
			showToolbar="true" showFormulabar="true" showContextMenu="true" showSheetbar="true"/>
	</div>
<script type="text/javascript">
	zk.afterMount(function(){
		//register client event on action component by jquery api 
		jq("#checkBtn").click(function(){
			postAjax("check");
		});
		jq("#resetBtn").click(function(){
			postAjax("reset");
		});
	});
	
	var ajaxUrl = "ajaxBook";//the servlet url to handle ajax request 
	
	function postAjax(action){
		//use zk client side api to get client side sparedsheet and desktop object 
		var zss = zk.Widget.$("$myzss");
		var desktop = zss.desktop;
		
		//Use jquery api to post ajax to your servlet (in this demo, it is AjaxBookServlet),
		//provide desktop id and spreadsheet uuid to access zk component data in your servlet 
		jq.ajax({url:ajaxUrl,
			data:{dtid:desktop.id,ssuuid:zss.uuid,action:action},
			dataType:'json'}).done(handleAjaxResult);
	}
	
	function handleAjaxResult(result){
		//eval any zkjs for updating zk ui
		if(result.zkjs){
			eval(result.zkjs);
		}
		
		//use your way handle your ajax action result
		if(result.action == "check" && result.validation){
			alert("validation passed");
		};
		
		//use your way to hanlde you ajax message or error
		if(result.message){
			alert(result.message);
		};
		if(result.error){
			alert(result.error);
		};
		
	}
	
	</script>	
</body>
</html>
