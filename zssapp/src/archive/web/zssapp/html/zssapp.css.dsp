<%@ page contentType="text/css;charset=UTF-8" %>
<%@ taglib uri="http://www.zkoss.org/dsp/web/core" prefix="c" %>

.clicked .z-toolbarbutton-cnt {
	background-image: url(${c:encodeURL('~./zssapp/image/toolbarbtn-ctr-clkd.gif')});
	background-position: 0 -500px;
	border: 1px solid #7F9DB9;
	padding: 1px;
}
.dpbtn-seld .z-dpbutton-cnt {
	border-bottom: 1px solid #7F9DB9;
	border-top: 1px solid #7F9DB9;
	border-left: 1px solid #7F9DB9;
	border-right: 1px solid #7F9DB9;
	background-color: #FCDE9A;
	padding: 1px 2px 1px 1px;
}
.z-dpbutton-dpover .z-dpbutton-cnt {
	border: 0;
	padding: 2px 3px 2px 2px;
}
.dpbtn-seld	div.z-dpbutton-cnt {
	background-image: none;
	background-color: #FCDE9A;
}
.dpbtn-seld .z-dpbutton-body {
	cursor: default;
}
.focusPosCombobox .z-combobox-rounded-btn {
	border:0px;
	background: url(${c:encodeURL('~./zssapp/image/combo-white.png')});
}
.focusPosCombobox .z-combobox-rounded-inp {
	background-image: none;
	font-size: 14px;
}
div.z-tabpanel { 
	border-bottom:none; 
	border-left:none;
	border-right:none; 
	padding:0; 
}

.toolbarSeparator {
	height: 30px; 
	width: 3px;
	background: url(${c:encodeURL('~./zssapp/image/sep.gif')}) repeat-y scroll 0 0 transparent;
}

.arialFont .z-menu-item-cnt, .arialFont .z-combo-item-text{
	font-size:14px;
	font-family:arial;
}
.calibriFont .z-menu-item-cnt, .calibriFont .z-combo-item-text{
	font-size:14px;
	font-family:calibri;
}
.timeFont .z-menu-item-cnt,.timeFont .z-combo-item-text{
	font-size:14px;
	font-family:time new roman;
}
.courierFont .z-menu-item-cnt,.courierFont .z-combo-item-text{
	font-size:14px;
	font-family:courier new;
}
.comicFont .z-menu-item-cnt,.comicFont .z-combo-item-text{
	font-size:14px;
	font-family:comic Sans MS;
}
.centuryFont .z-menu-item-cnt,.centuryFont .z-combo-item-text{
	font-size:14px;
	font-family:century gothic;
}
.georgiaFont .z-menu-item-cnt,.georgiaFont .z-combo-item-text{
	font-size:14px;
	font-family:georgia;
}
.impactFont .z-menu-item-cnt,.impactFont .z-combo-item-text{
	font-size:14px;
	font-family:impact;
}
.toolbarMask{
	position:absolute;
	visibility: hidden;
	left: 0;
	background:none repeat scroll 0 0 #E0E1E3;
	z-index:100; 
	opacity: 0.6;
	-moz-opacity: .6;
	filter: alpha(opacity=60);
}
.zssapp-menubar {
	border-top : 1px solid #D8D8D8;
	border-bottom : 1px solid #D8D8D8;
	background-image: url(${c:encodeURL('~./zssapp/image/menu-bg.png')});
}
.zssapp-menubar .z-menubar-hor {
	border-top : 0;
	border-bottom : 0;
}
.zssapp-empty {
	border: 1px solid #D8D8D8;
	background-color: #FEFEFE;
	background-image: -moz-linear-gradient(top, #FEFEFE, #E4ECF7); /* FF3.6 */
	background-image: -webkit-gradient(linear,left top,left bottom,color-stop(0, #FEFEFE),color-stop(1, #E4ECF7)); /* Saf4+, Chrome */
    filter:  progid:DXImageTransform.Microsoft.gradient(startColorStr='#FEFEFE', EndColorStr='#E4ECF7'); /* IE6,IE7 */
    -ms-filter: "progid:DXImageTransform.Microsoft.gradient(startColorStr='#FEFEFE', EndColorStr='#E4ECF7')"; /* IE8 */
}
.zssapp-tabs .z-tab-hl, .zssapp-tabs .z-tab-hr, .zssapp-tabs .z-tab-hm {
	background: #F9F9F9;
}
.zssapp-tabs .z-tab {
	border: 1px solid #D4D4D4;
	border-top: 0;
}
.zssapp-tabs .z-tab-seld, 
.zssapp-tabs .z-tab-seld .z-tab-hl, 
.zssapp-tabs .z-tab-seld .z-tab-hr, 
.zssapp-tabs .z-tab-seld .z-tab-hm {
	background-color: #FFFFFF;
}

.fastIconWin .z-window-popup-cnt-noborder{
	background: #f6f5f5 none;
	border:1px solid #C5C5C5;
}

.autoFilterListbox .z-listitem-seld, .autoFilterListbox .z-listitem-seld {
	background-image: none;
}
.autoFilterListbox .z-listitem-over {
	background-image: none;
	background: #c5e8fa;
}
.redItem .z-listcell-cnt {
	color: red;
}
.fileDeleteWin .z-window-embedded-cnt { 
	padding:7px; 
}
.fileDeleteWin .z-list-cell-cnt{
	padding:5px 5px 5px 2px;
}
.fileDeleteWin .hintLabel{
	margin:5px;
	font-size:16px;
	font-family:Calibri;
}
.fileExportWin .z-window-embedded-cnt { 
	padding:7px; 
}
.fileExportWin .z-list-cell-cnt{
	padding:5px 5px 5px 2px;
}
.fileExportWin .hintLabel{
	margin:5px;
	font-size:16px;
	font-family:Calibri;
}
.fileOpenViewWin .hintLabel{
	font-family:Calibri;
}
.z-grid-body tr.btnRow td.z-row-inner {
		background: none;
}