<%@page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@taglib prefix="zss" uri="http://www.zkoss.org/jsp/zss"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
		<title>Application for Leave</title>
		<zss:zkhead/>
	</head>
<body>
	<div>
		<zss:spreadsheet id="myzss" src="/WEB-INF/books/application_for_leave.xlsx"
			width="800px" height="600px" 
			maxrows="100" maxcolumns="20"
			showToolbar="true" showFormulabar="true" showContextMenu="true" showSheetbar="true"/>
	</div>
</body>
</html>
