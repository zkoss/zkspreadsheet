<%@ page contentType="text/css;charset=UTF-8" %>
<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>
<c:include page="~./zss/css/ss.css.dsp"/>
.zscell {
	display: inline-block;/* DIFF */
}
.zscell-over {
	border-right: 0px;/* DIFF */
	margin-right: 1px;/* DIFF */
}
.zscell-inner-mid-left {
	display: -webkit-box;
	-webkit-box-orient: horizontal;
	-webkit-box-pack: start;
	-webkit-box-align: center;
}

.zscell-inner-mid-center {
	display: -webkit-box;
	-webkit-box-orient: horizontal;
	-webkit-box-pack: center;
	-webkit-box-align: center;
}

.zscell-inner-mid-right {
	display: -webkit-box;
	-webkit-box-orient: horizontal;
	-webkit-box-pack: end;
	-webkit-box-align: center;
}

.zscell-inner-btm-left {
	display: -webkit-box;
	-webkit-box-orient: horizontal;
	-webkit-box-pack: start;
	-webkit-box-align: end;
}

.zscell-inner-btm-center {
	display: -webkit-box;
	-webkit-box-orient: horizontal;
	-webkit-box-pack: center;
	-webkit-box-align: end;
}

.zscell-inner-btm-right {
	display: -webkit-box;
	-webkit-box-orient: horizontal;
	-webkit-box-pack: end;
	-webkit-box-align: end;
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