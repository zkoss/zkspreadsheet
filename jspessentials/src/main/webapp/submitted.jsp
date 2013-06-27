<%@page import="java.util.Date"%>
<%@page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@taglib prefix="zss" uri="http://www.zkoss.org/jsp/zss"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
	<title>Application for Leave</title>
	</head>
<%
	Date from = new Date(Long.parseLong(request.getParameter("from")));
	Date to = new Date(Long.parseLong(request.getParameter("to")));
	int total = Integer.parseInt(request.getParameter("total"));//should count by from and to in real case
	String reason = request.getParameter("reason");
	
	String applicant = request.getParameter("applicant");
	Date requestDate = new Date(Long.parseLong(request.getParameter("requestDate")));
	String archive = request.getParameter("archive");
%>
<body>
	<table>
		<tr>
			<td colspan="2">Processed Data</td>
		</tr><tr>
			<td>Request Date:</td><td><%=requestDate%></td>
		</tr><tr>
			<td>From:</td><td><%=from%></td>
		</tr><tr>
			<td>To:</td><td><%=to%></td>
		</tr><tr>
			<td>Total Days:</td><td><%=total%></td>
		</tr><tr>
			<td>Reason:</td><td><%=reason%></td>
		</tr><tr>
			<td>Applicant:</td><td><%=applicant%></td>
		</tr><tr>
			<td>Archive:</td><td><%=archive%></td>
		</tr>
	</table>
	
</body>
</html>
