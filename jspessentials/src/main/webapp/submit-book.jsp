<%@page import="org.zkoss.zk.ui.Execution"%>
<%@page import="org.zkoss.zssex.ui.ExecutionBridge"%>
<%@page import="org.zkoss.zss.api.Importers"%>
<%@page import="org.zkoss.zss.api.model.Book"%>
<%@page import="java.net.URL"%>
<%@page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@taglib prefix="zss" uri="http://www.zkoss.org/jsp/zss"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<zss:zkhead/>
<title>My First ZK Spreadsheet JSP application</title>
</head>
<%
	//prevent page cache in browser side
	response.setHeader("Pragma", "no-cache"); 
	response.setHeader("Cache-Control", "no-store, no-cache"); 

	
	ExecutionBridge eb;
	new ExecutionBridge(getServletContext(),request,response){
		protected void process(Execution execution) throws Exception{
			Book book = (Book)_request.getSession().getAttribute("book");
			if(book==null){
				URL bookUrl = getServletContext().getResource("/WEB-INF/books/sample.xlsx");
				book = Importers.getImporter().imports(bookUrl, "mybook");
				_request.getSession().setAttribute("book",book);
				book.setShareScope("session"); //will get java.lang.IllegalStateException: Not in an execution
			}	
		}
	}.process();
%>
<body>
	<form action="submitted-book.jsp" >
		<button>Submit</button>
	</form>
	<div width="100%" height="100%">
		<zss:spreadsheet id="myzss" book="${book}"
			width="800px" height="600px" 
			maxrows="200" maxcolumns="40"
			showToolbar="true" showFormulabar="true" showContextMenu="true" showSheetbar="true"/>
	</div>
</body>
</html>