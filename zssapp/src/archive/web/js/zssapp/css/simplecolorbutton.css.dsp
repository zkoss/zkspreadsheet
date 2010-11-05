<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>
<%-- Colorbox --%>
.z-colorbtn, .z-colorbtn-currcolor {
	width: 16px;
}
.z-colorbtn {
	border: 1px solid transparent;
	-moz-border-radius: 3px;
	-webkit-border-radius: 3px;
	margin:0 2px;
	overflow: hidden;
	display: inline-block;
	vertical-align: middle;
	padding: 3px;
	position: relative;
	<c:if test="${c:isExplorer()}">
		zoom: 1;
		*display: inline;
	</c:if>
}
.z-colorbtn-over {
	background: #E3F9FF;
	border: 1px solid #7EAAC6;
}
.z-colorbtn-btn, .z-colorbtn-currcolor {
	overflow: hidden;
	cursor: pointer;
	position: absolute;
}
.z-colorbtn-btn {
	bottom: 1px;
	padding-bottom: 1px;
}
.z-colorbtn-currcolor {
	bottom: 1px;
	left: 4px;
	height: 3px;
	z-index: 8;
}
.z-colorpicker, .z-colorpicker-gradient, z-colorpicker-bar {
	border: 1px solid #86A4BE;
}
.z-colorbtn-disb {
	opacity: .6;
	-moz-opacity: .6;
	filter: alpha(opacity=60);
}
.z-colorbtn-disb, .z-colorbtn-disb * {
	cursor: default !important;
	color: #AAA !important;
}
<%-- Popup shadow --%>
.z-colorbtn-palette-pp, .z-colorbtn-picker-pp {
	-moz-box-shadow: 1px 1px 3px rgba(0, 0, 0, 0.5);
	-webkit-box-shadow: 1px 1px 3px rgba(0, 0, 0, 0.5);
}
<%-- Colorpicker --%>
.z-colorpicker {
	width: 380px;
	height: 310px;
	overflow: hidden;
	position: relative;
	background-color:#FFFFFF;
}
<%-- Colorpicker gradient--%>
.z-colorpicker-gradient, .z-colorpicker-overlay {
	width: 256px;
	height: 256px;
}
.z-colorpicker-gradient {
	left: 7px;
	top: 9px;
	position: absolute;
	cursor: crosshair;
}
.z-colorpicker-overlay {
	background-image: url(${c:encodeURL('~./zkex/img/colorbox/colorpicker_gradient.png')});
}
.z-colorpicker-circle {
	position: absolute;
	top: 0;
	left: 0;
	width: 11px;
	height: 11px;
	overflow: hidden;
	background-image: url(${c:encodeURL('~./zkex/img/colorbox/colorpicker_select.gif')});
	margin: -5px 0 0 -5px;
}
<%-- Colorpicker hue--%>
.z-colorpicker-hue {
	position: absolute;
	top: 9px;
	left: 272px;
	width: 27px;
	height: 256px;
}
.z-colorpicker-bar, .z-colorpicker-arrows {
	position: absolute;
	overflow: hidden;
	cursor: n-resize;
}
.z-colorpicker-bar {
	width: 12px;
	height: 256px;
	left: 7px;
	background-image: url(${c:encodeURL('~./zkex/img/colorbox/colorpicker_hue.png')});
}
.z-colorpicker-arrows {
	width: 27px;
	height: 9px;
	background-image: url(${c:encodeURL('~./zkex/img/colorbox/colorpicker_arrows.gif')});
	margin: -4px 0 0 0;
	left: 0;
}
<%-- Colorpicker colors --%>
.z-colorpicker-color {
	top: 12px;
	left: 315px;
	border: double;
	position: absolute;
	background: transparent;
}
.z-colorpicker-newcolor, .z-colorpicker-currcolor {
	position: relative;
	width: 48px;
	height: 32px;
}
.z-colorpicker-newcolor {
	border-bottom: 1px solid;
}
.z-colorpicker-currcolor {
	border-top: 1px solid;
}
<%-- Colorpicker input --%>
.z-colorpicker-r, .z-colorpicker-g, .z-colorpicker-b,
.z-colorpicker-h, .z-colorpicker-s, .z-colorpicker-v {
	position: absolute;
	left: 310px;
}
.z-colorpicker-r {
	top: 100px;
}
.z-colorpicker-g {
	top: 125px;
}
.z-colorpicker-b {
	top: 150px;
}
.z-colorpicker-h {
	top: 190px;
}
.z-colorpicker-s {
	top: 215px;
}
.z-colorpicker-v {
	top: 240px;
}
.z-colorpicker-r-text, .z-colorpicker-g-text, .z-colorpicker-b-text,
.z-colorpicker-h-text, .z-colorpicker-s-text, .z-colorpicker-v-text,
.z-colorpicker-r-inp, .z-colorpicker-g-inp, .z-colorpicker-b-inp,
.z-colorpicker-h-inp, .z-colorpicker-s-inp, .z-colorpicker-v-inp,
.z-colorpicker-hex-text, .z-colorpicker-hex-inp {
	font-family: ${fontFamilyC};
	font-size: ${fontSizeM};
	font-weight: normal;
}
.z-colorpicker-r-inp, .z-colorpicker-g-inp, .z-colorpicker-b-inp,
.z-colorpicker-h-inp, .z-colorpicker-s-inp, .z-colorpicker-v-inp {
	margin-left: 5px;
	text-align: left;
}
<%-- Colorpicker input : hex --%>
.z-colorpicker-hex {
	position: absolute;
	top: 275px;
	left: 10px;
}
.z-colorpicker-hex-inp {
	margin-left: 5px;
}
<%-- Colorpicker btn --%>
.z-colorpicker-ok-btn, .z-colorpicker-cancel-btn {
	cursor: pointer;
	top: 275px;
	position: absolute;
}
.z-colorpicker-ok-btn {
	left: 328px;
	width: 42px;
}
.z-colorpicker-cancel-btn {
	left: 275px;
	width: 42px;
}
<%-- Color Palette --%>
.z-colorpalette {
	width: 256px;
	height: 220px;
	border: 1px solid #86A4BE;
	background-color: #FFFFFF;
}
.z-colorpalette-newcolor {
	width: 50px;
	height: 20px;
	border: 1px solid #86A4BE;
	margin: 4px 2px;
	position: relative;
	left: 120px;
}
.z-colorpalette-hex-inp, .z-colorpalette-btn {
	position: absolute;
	top: 5px;
	left: 180px;
}
.z-colorpalette-btn {
	left: 200px;
	width: 20px;
	height: 20px;
	cursor: pointer;
	background: url(${c:encodeURL('~./zkex/img/colorbox/colorpicker-btn.gif')});
}
.z-colorpalette-colorbox {
	width: 12px;
	height: 12px;
	float: left;
	border: 1px solid #FFFFFF;
	display: inline;
}
.z-colorpalette-colorbox-over {
	border: 1px solid #000000;
}
.z-colorbtn-palette-btn,
.z-colorbtn-picker-btn {
	z-index: 10;
	width: 22px;
	height: 22px;
	cursor: pointer;
	position: absolute;
	background: url(${c:encodeURL('~./zkex/img/breeze/colorbox/cb-buttons.gif')});
}
.z-colorbtn-palette-btn {
	left: 6px;
	top: 5px;
	background-position:  0 0;
}
.z-colorbtn-picker-btn {
	left: 31px;
	top: 5px;
	background-position:  0 -44px;
}
.z-palette-btn .z-colorbtn-palette-btn {
	left: 6px;
	background-position:  0 -22px;
}
.z-palette-btn .z-colorbtn-picker-btn {
	left: 31px;
}
.z-picker-btn .z-colorbtn-picker-btn {
	background-position:  0 -66px;
}
<%-- Color cell --%>
.z-colorbtn-pp {
	border: 1px solid #86A4BE;
	position: absolute;
	width: 180px;
	height: 120px;
	background: #FFF;
}
.z-colorbtn-pp .z-colorbtn-cell {
	display: inline;
	float: left;
	height: 18px;
	overflow: hidden;
	padding: 2px;
	width: 18px;
	pointer: cursor;
}
.z-colorbtn-panel {
	padding: 2px;
}
.z-colorbtn-pp .z-colorbtn-cell-over {
	padding: 1px;
	border: 1px solid #666666;
}
.z-colorbtn-pp .z-colorbtn-cell-cnt {
	border: 1px solid #808080;
	height: 10px;
	overflow: hidden;
	width: 11px;
	margin: 2px;
}