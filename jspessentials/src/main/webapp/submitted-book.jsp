<%@page import="org.zkoss.zss.api.model.Sheet"%>
<%@page import="org.zkoss.zss.api.Ranges"%>
<%@page import="org.zkoss.zss.api.Range"%>
<%@page import="org.zkoss.zss.api.Importers"%>
<%@page import="org.zkoss.zss.api.model.Book"%>
<%@page import="java.net.URL"%>
<%@page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@taglib prefix="zss" uri="http://www.zkoss.org/jsp/zss"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>My First ZK Spreadsheet JSP application</title>
</head>
<body>
<%
	Book book = (Book)session.getAttribute("book");
	if(book==null){
		out.println("book data not found");
		return;
	}
	
	
	Sheet sheet = book.getSheetAt(0);
	
	Range range = Ranges.range(sheet,"D4");
	String text = range.getCellFormatText();
	Double value = (Double)range.getCellValue();
	
	out.println("value is "+value+", text is "+text);
	
	//if the processing is done, you could consider to remove it for session.
	session.removeAttribute("book");
			
	//TODO how to clear a book if not in execution, it refer to a event queue
%>
</body>
</html>