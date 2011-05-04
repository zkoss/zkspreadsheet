<%@ page contentType="text/css;charset=UTF-8" %>
<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>
<c:include page="~./zss/css/ss.css.dsp"/>
/*
To compete zk 3.0.6, in IE, i must add 'div.' before rule, if there is font-size attribute 

*/
.zscell {
	display: inline-block; /* DIFF */
	zoom: 1;
	*display: inline;
}
.zscell-over {
	border-right: 0px;/* DIFF */
	margin-right: 1px;/* DIFF */
}

.zsrow{
	font-size:0;/* bug 1990408 */
}

<c:if test="${c:isExplorer() && !c:isExplorer7()}">
.zsrow {
	position: static;
	overflow-y: hidden;
}
</c:if>

.zsleftcelltxt {
	font-size: 10px;/*  bug 1990408,to correct zsrow font-size 0*/
}

<c:if test="${c:isExplorer() && !c:isExplorer7()}">
.zstopblock {
	top: -1px;
}
</c:if>
<c:if test="${c:isExplorer7()}">
.zstopblock {
	top: 1px;
}
</c:if>

.zstopcell {
	display: inline-block; /* DIFF */
	zoom: 1;
	*display: inline;
}
.zscorner {
	*font-size: 0; /* DIFF */
}
.zsfocmark {
	font-size: 0; /* DIFF */
}
.zsfocmarki{/*this is created for IE*/
	font-size: 0; /* DIFF */
	filter: alpha(opacity=0);/* DIFF */
	background: #FFFFFF;/* DIFF */
}
.zsselect {
	font-size: 0; /* DIFF */
}
.zsselecti{/*this is created for IE*/
	font-size: 0; /* DIFF */
	filter: alpha(opacity=0);/* DIFF */
	background: #FFFFFF;/* DIFF */
}
.zsselecti-r{
	filter: alpha(opacity=50);/* DIFF */
	background-color: #E3ECF7;
}
.zsselchgi{
	filter: alpha(opacity=40);/* DIFF */
	background-color: #99FFAA;
}

.zsseldot {
	font-size: 0; /* DIFF */
}
.zshboun{
	display: inline-block; /* DIFF */
	zoom: 1;
	*display: inline;
}
.zshbouni{
	*position: absolute; /* DIFF */
	display: inline-block; /* DIFF */
	zoom: 1;
	*display: inline;
}

.zshbounw{
	display: inline-block; /* DIFF */
	zoom: 1;
	*display: inline;
	*position: absolute; /* DIFF */
	*left:7px;
}

.zsvbouni{
	*position: absolute;/* DIFF */
	font-size: 0; /* DIFF */
}

.zsvbounw{
	font-size: 0; /* DIFF */
	*position: absolute; /* DIFF */
	*top:7px;
}

.zsscroll {
	background-color:#CAD7E6;
}
.zsdata {
	background-color:#EEEEFF;
}

.zsblock {
	background-color:#FFFFFF;
}

.zssmask{
	filter:alpha(opacity=75);
}
.zssmask2{
	#position:absolute;
	#top:50%;
}
.zssmasktxt{
	#position:relative;
	#top:-50%"
}
