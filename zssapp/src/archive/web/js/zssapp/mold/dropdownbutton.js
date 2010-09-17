/* colorbox.js
 Purpose:
 
 Description:
 
 History:
 Wed Nov  4 12:25:37 TST 2009, Created by sam
 Copyright (C) 2009 Potix Corporation. All Rights Reserved.
 
 */
function (out) {
	var zcls = this.getZclass(),
		popup = this.popup;
	out.push('<div', this.domAttrs_(), '><div class="',
		zcls, '-body"><div ', this.domTextStyleAttr_(), 
		'class="', zcls, '-cnt" id="', this.uuid, '-cnt">', this.domContent_());
	if (popup)
		popup.redraw(out);
	out.push('</div><div id="', this.uuid, '-btn" class="', zcls, '-btn"><div class="', zcls, '-arrow"></div></div></div></div>');
}