<%@ page contentType="text/css;charset=UTF-8" %>
<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>
<c:include page="~./zss/css/ss.css.dsp"/>
.zscell-inner-mid-left {
	display: -moz-box;
	-moz-box-orient: horizontal;
	-moz-box-pack: start;
	-moz-box-align: center;
}
.zscell-inner-mid-center {
	display: -moz-box;
	-moz-box-orient: horizontal;
	-moz-box-pack: center;
	-moz-box-align: center;
}

.zscell-inner-mid-right {
	display: -moz-box;
	-moz-box-orient: horizontal;
	-moz-box-pack: end;
	-moz-box-align: center;
}

.zscell-inner-btm-left {
	display: -moz-box;
	-moz-box-orient: horizontal;
	-moz-box-pack: start;
	-moz-box-align: end;
}

.zscell-inner-btm-center {
	display: -moz-box;
	-moz-box-orient: horizontal;
	-moz-box-pack: center;
	-moz-box-align: end;
}

.zscell-inner-btm-right {
	display: -moz-box;
	-moz-box-orient: horizontal;
	-moz-box-pack: end;
	-moz-box-align: end;
}
.zsscroll {
	background-color:#CAD7E6;
}
.zsdata {
	background-color:#EEEEFF;
}
<c:if test="${c:browser('ie')}">
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

.zsleftcelltxt {
	font-size: 10px;/*  bug 1990408,to correct zsrow font-size 0*/
}

<c:if test="${c:isExplorer() && !c:isExplorer7()}">
.zstopblock {
	top: -1px;
}
</c:if>

.zscell-inner-mid-left {
	display: -moz-box;
	-moz-box-orient: horizontal;
	-moz-box-pack: start;
	-moz-box-align: center;
}
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
</c:if>
<c:if test="${c:isSafari()}">
.zscell {
	display: inline-block;/* DIFF */
}
.zscell-over {
	border-right: 0px;/* DIFF */
	margin-right: 1px;/* DIFF */
}

.zstopblock {
	top: 1px;
}

.zstopcell {
	display: inline-block;/* DIFF */
}
.zshboun{/*hboundary*/
	display: inline-block; /* DIFF */
}

.zshbouni{
	display: inline-block; /* DIFF */
}

.zshbounw{
	display: inline-block; /* DIFF */
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
</c:if>