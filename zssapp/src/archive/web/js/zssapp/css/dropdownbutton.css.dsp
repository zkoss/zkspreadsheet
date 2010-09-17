<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>
<%-- dpbutton modify from Toolbarbutton --%>
.z-dpbutton {
	display:-moz-inline-box;
	display: inline-block;
	position: relative;
	cursor: pointer;
	margin: 0 2px;
	vertical-align: middle;	
	padding: 1px 0;
	zoom: 1;
}
.z-dpbutton-over {
	border-top: 1px solid #7EAAC6;
	border-bottom: 1px solid #7EAAC6;
	padding: 0;
}
.z-dpbutton-body {
	position: relative;
	margin: 0 -1px;
	padding: 0 1px;
	vertical-align: middle;:
	zoom: 1;
}
.z-dpbutton-over .z-dpbutton-body {
	border-left: 1px solid #7EAAC6;
	border-right: 1px solid #7EAAC6;
	padding: 0;
}
<%-- MODIFY: remove font-size --%>
.z-dpbutton-cnt {
	padding: 2px 2px;
	position: relative;
	zoom: 1;
	font-family: ${fontFamilyT};
	font-weight: normal;
}

<c:if test="${c:isExplorer() and not c:browser('ie8')}">
.z-dpbutton {
	display: inline;
}

<c:if test="${c:browser('ie6-')}">
.z-dpbutton,
.z-dpbutton-body,
.z-dpbutton-cnt {
	display: inline;
	position: relative;
}
.z-dpbutton-body {
	float: left;
}
</c:if>
</c:if>
.z-dpbutton-disd * {
	color:gray !important;
	cursor:default !important;
}
.z-dpbutton-disd ${c:isExplorer() ? '*': ''} { <%-- bug 3022237 --%>
	opacity: .5;
	-moz-opacity: .5;
	filter: alpha(opacity=50);
}

<%-- ADDED --%>
.z-dpbutton-body {
	height: 20px;
}
.z-dpbutton-cnt, .z-dpbutton-btn {
	margin:	0;
	border: 0;
	line-height: 0;
	font-size: 0;
	top: 0;
	display: inline-block;
	<c:if test="${c:isExplorer()}">
		zoom: 1;
		*display: inline;
	</c:if>
}
.z-dpbutton-cnt {
	padding-right: 3px;
	vertical-align: middle;
}
.z-dpbutton-btn {
	position: relative;
	padding: 2px 2px 2px 0;
	vertical-align: middle;
	width: 9px;
	height: 16px;
}
.z-dpbutton-btn .z-dpbutton-arrow {
	width: 9px;
	height: 16px;
	background: url(${c:encodeURL('~./zssapp/image/dp-arrow.gif')}) no-repeat;
}
<%-- Main area no clickable --%>
.z-dpbutton-noclk {
	cursor: default;
}
.z-dpbutton-noclk  .z-dpbutton-btn {
	cursor: pointer;
}
<%-- Main area mouse over --%>
.z-dpbutton-over .z-dpbutton-cnt {
	background: url(${c:encodeURL('~./zssapp/image/dp-over.png')}) repeat-x;
}
.z-dpbutton-over .z-dpbutton-btn {
	background-position: 0 -500px;
	background-image:url(${c:encodeURL('~./zul/img/breeze/button/toolbarbtn-ctr.gif')});
}
<%-- Doprdown area mouse over --%>
.z-dpbutton-dpover {
	border-top: 1px solid #7EAAC6;
	border-bottom: 1px solid #7EAAC6;
	padding: 0;
}
.z-dpbutton-dpover .z-dpbutton-body {
	border-left: 1px solid #7EAAC6;
	border-right: 1px solid #7EAAC6;
	padding: 0;
}
.z-dpbutton-dpover .z-dpbutton-btn {
	background: url(${c:encodeURL('~./zssapp/image/dp-over.png')}) repeat-x;
}
.z-dpbutton-dpover .z-dpbutton-cnt {
	background-position: 0 -500px;
	background-image:url(${c:encodeURL('~./zul/img/breeze/button/toolbarbtn-ctr.gif')});
}
<c:if test="${c:isExplorer() and not c:browser('ie8')}">
.z-dpbutton-cnt, .z-dpbutton-btn {
	top: -1px;
}
</c:if>