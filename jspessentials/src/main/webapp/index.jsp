<%@page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@taglib prefix="zssjsp" uri="http://www.zkoss.org/jsp/zss"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<title>Application for Leave</title>
		<zssjsp:head/>
	</head>
<body>
	<div>
		<zssjsp:spreadsheet id="myzss" src="/WEB-INF/books/application_for_leave.xlsx"
			width="1024px" height="768px" 
			maxrows="100" maxcolumns="20"
			showToolbar="true" showFormulabar="true" showContextMenu="true" showSheetbar="true"/>
	</div>
</body>
</html>
