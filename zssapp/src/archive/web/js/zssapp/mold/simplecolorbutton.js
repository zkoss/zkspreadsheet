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
		cols = this._colors,
		colSize = this._colors.length;

	out.push('<div', this.domAttrs_(), '>',
		'<i id="', uuid, '-currcolor" class="', zcls, '-currcolor"></i>',
		'<img id="', uuid, '-btn" class="', zcls, '-btn"></img>',
		'<div id="', uuid, '-pp" style="display:none;" class="', zcls, '-pp" >',
		'<div class="', zcls,'-panel">');
	for (var i = 0; i < colSize; i++) {
		out. push('<div class="', zcls, '-cell">',
				'<div style="background: ', cols[i], ';" class="', zcls, '-cell-cnt"><i style="display: none">',
				cols[i],'</i></div></div>');
	}
	out.push('</div></div></div>');
}