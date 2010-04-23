/* colorbox.js
 Purpose:
 
 Description:
 
 History:
 Wed Nov  4 12:25:37 TST 2009, Created by sam
 Copyright (C) 2009 Potix Corporation. All Rights Reserved.
 
 */
function (out) {
	var zcls = this.getZclass(),
		uuid = this.uuid,
		imgSrc = this.getImage(),
		picker = this._picker,
		palette = this._palette;
	out.push('<div', this.domAttrs_(), '>',
	'<i id="', uuid, '-currcolor" class="', zcls, '-currcolor"></i>',
	'<img id="', uuid, '-btn" class="', zcls, '-btn" src="', imgSrc != null ? imgSrc : '', '"></img>');
	if (picker) {
		out.push('<div id="', uuid, '-picker-pp" class="', zcls, '-picker-pp" style="display:none">');
		picker.redraw(out);
		out.push('</div>');
	}
	if (palette) {
		out.push('<div id="', uuid, '-palette-pp" class="', zcls, '-palette-pp" style="display:none">');
		palette.redraw(out);
		out.push('</div>');
	}	
	out.push('</div>');
}