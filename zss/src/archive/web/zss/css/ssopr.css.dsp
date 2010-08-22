<%@ page contentType="text/css;charset=UTF-8" %>
<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>
<c:include page="~./zss/css/ss.css.dsp"/>
.zscell {
	display: inline-block;/* DIFF */
	vertical-align:text-top;
	/*vertical-align: top;*//* DIFF : important, vertical align will cause opera scrollbar disappear*/
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