/* Editbox.js

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mon Apr 23, 2007 17:29:18 AM , Created by sam
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/


/**
 * 
 * Editbox represent a edit area for cell
 */
zss.Editbox = zk.$extends(zk.Object, {
	$init: function (sheet) {
		this.$supers('$init', arguments);
		var wgt = sheet._wgt;
			editbox = wgt.$n('eb');
		this.id = editbox.id;
		this.sheetid = sheet.sheetid;
		this.sheet = sheet;
		this.comp = editbox;
		this.comp.ctrl = this;
		this.row = -1;
		this.col = -1;
		this.disabled = true;

		wgt.domListen_(editbox, 'onBlur', '_doEditboxBlur');
		wgt.domListen_(editbox, 'onKeyPress', '_doEditboxKeyPress');
		wgt.domListen_(editbox, 'onKeyDown', '_doEditboxKeyDown');
		wgt.domListen_(editbox, 'onKeyUp', '_doEditboxKeyUp');
	},
	cleanup: function () {
		var wgt = this.sheet._wgt,
			editbox = wgt.$n('eb');
		wgt.domUnlisten_(editbox, 'onBlur', '_doEditboxBlur');
		wgt.domUnlisten_(editbox, 'onKeyPress', '_doEditboxKeyPress');
		wgt.domUnlisten_(editbox, 'onKeyDown', '_doEditboxKeyDown');
		wgt.domUnlisten_(editbox, 'onKeyUp', '_doEditboxKeyUp');
		
		this.invalid = true;

		if (this.comp) this.comp.ctrl = null;
		this.comp = this.sheet = null;
	},
	_doKeydown: function (evt) {
		var keycode = evt.keyCode;
		switch (keycode) {
		case 35: //End
			if (evt.altKey) {
				if (this.col + this.sw < this.sheet.maxCols - 1) {
					this.sw++;
					this.adjust("w");
				}
				evt.stop();
			}
			break;
		case 34: //PgDn
			if (evt.altKey) {
				if(this.row + this.sh <  this.sheet.maxRows - 1) {
					this.sh++;
					this.adjust("h");
				}
				evt.stop();
			}
			break;
		case 36: //Home
			if (evt.altKey) {
				if (this.sw > 0) {
					this.sw--;
					this.adjust("w");
				}
				evt.stop();
			}
			break;
		case 33: //PgUp
			if (evt.altKey) {
				if (this.sh > 0) {
					this.sh--;
					this.adjust("h");
				}
				evt.stop();
			}
			break;
		}
	},
	_doKeyup: function (evt, val) {
		var sheet = this.sheet,
			value = sheet.editor.currentValue(),
			keycode = evt.keyCode;

		sheet._wgt.fire('onEditboxEditing', {token: '', sheetId: sheet.serverSheetId, clienttxt: value});
	},
	/**
	 * Sets edit box disabled
	 * @param boolean disabled
	 */
	disable: function (disabled) {
		this.comp.style.backgroundColor = disabled ? "#DDDDDD" : "#EFECFF";
		this.disabled = disabled;
	},
	/**
	 * 
	 */
	edit: function (cellcmp, row, col, value) {//cmp is the focused cell
		this.disable(false);
		this.row = row;
		this.col = col;
		this.sw = 0;//width to show
		this.sh = 0;//height to show
		var sheet = this.sheet,
			txtcmp = cellcmp.lastChild,
			editorcmp = this.comp,
			$edit = jq(editorcmp);

		editorcmp.value = value;
		var w = cellcmp.ctrl.overflowed ? (cellcmp.firstChild.offsetWidth + this.sheet.cellPad) : (cellcmp.offsetWidth),
			h = cellcmp.offsetHeight,
			l = cellcmp.offsetLeft,
			t = cellcmp.parentNode.offsetTop;
			blockcmp = cellcmp.ctrl.block.comp;//add block offset

		//IE6 only: .zsrow css position changed from "relative" -> "static" for cell vertical merge which cause cellcmp.offsetLeft change
		l += zk.ie6_ ? 0 : blockcmp.offsetLeft;//block 
		t += blockcmp.offsetTop;//block
		t -= 1;//position adjust
		w -= 1;
		h -= 1;
		
		if (zk.ie || zk.safari || zk.opera)
			//the display in different browser. 
			w -= 2;

		this.editingWidth = w;
		this.editingHeight = h;

		//issue 228: firefox need set display block, but IE can not set this.
		$edit.css({'width': jq.px0(w), 'height': jq.px0(h), 'left': jq.px(l), 'top': jq.px(t), 'line-height': jq.px0(sheet.lineHeight)});
		if (!zk.ie)
			$edit.css('display', 'block');

		zcss.copyStyle(txtcmp, editorcmp, ["font-family","font-size","font-weight","font-style","color","text-decoration","text-align"],true);
		zcss.copyStyle(cellcmp, editorcmp, ["background-color"], true);

		//move text cursor position to last
		fun = function () {
			var pos = editorcmp.value.length;
			if (editorcmp.setSelectionRange) {
				editorcmp.setSelectionRange(pos,pos);
			} else if (editorcmp.createTextRange) {
				var range = editorcmp.createTextRange();
				range.move('character', pos);
				range.select();
			}
		};

		if (!zk.safari && !zk.ie) fun();//safari must run after timeout
		setTimeout(function(){
			//issue 228: ie focus event need after show
			zk.ie ? $edit.show().focus() : editorcmp.focus();
			//issue 230: IE cursor position is not at the text end when press F2
			if (zk.safari || zk.ie) fun();
		}, 25);
		this.autoAdjust(true);
	},
	cancel: function () {
		this.disable(true);
		jq(this.comp).val('').css('display', 'none');
		this.row = this.col = -1;
	},
	stop: function () {
		this.disable(true);
		var editorcmp = this.comp,
			str = editorcmp.value;
		jq(editorcmp).val('').css('display', 'none');
		this.row = this.col = -1;
		return str;
	},
	currentValue: function () {
		return this.comp.value;
	},
	newLine: function () {
		var editorcmp = this.comp,
			sel = zk(editorcmp).getSelectionRange();	
		if (sel[0] > sel[1]) {
			var t = sel[1];
			sel[1] = sel[0];
			sel[0] = t;
		}
		var str = editorcmp.value,
			len = str.length,
			str1 = str.substring(0, sel[0]); 
		str = str1 +'\n'+ str.substring(sel[1], len);
		this.comp.value = str;
		
		var pos = sel[0];
		//move text cursor position to last
		fun = function () {
			if (editorcmp.setSelectionRange) { 
				pos = zk.opera ? pos + 2 : pos + 1;
				editorcmp.setSelectionRange(pos, pos);
			} else if (editorcmp.createTextRange) {//IE
				var range = editorcmp.createTextRange(),
					i = -1;
				pos = pos + 1;
				while((i = str1.indexOf('\n',i+1)) >= 0) {
					pos--;
				}
				range.move('character', pos);
				range.select();
			}
		};
		fun();
		this.autoAdjust();
	},
	adjust: function (type) {
		var editorcmp = this.comp;

		if (type == "w") {
			var custColWidth = this.sheet.custColWidth,
				w = custColWidth.getStartPixel(this.col + this.sw + 1) - custColWidth.getStartPixel(this.col);
			if (zk.ie || zk.safari || zk.opera)
				w -= 2;
			jq(editorcmp).css('width', jq.px0(w));
		} else if (type=="h") {
			var custRowHeight = this.sheet.custRowHeight,
				h = custRowHeight.getStartPixel(this.row + this.sh + 1) - custRowHeight.getStartPixel(this.row);
			jq(editorcmp).css('height', jq.px0(h));
		}
	},
	autoAdjust: function (forceadj) {
		var local = this;
		setTimeout(function() {
			var editorcmp = local.comp,
				ch = editorcmp.clientHeight,
				cw = editorcmp.clientWidth,
				sw = editorcmp.scrollWidth,
				sh = editorcmp.scrollHeight,
				hsb = zkS._hasScrollBar(editorcmp),//has hor scrollbar
				vsb = zkS._hasScrollBar(editorcmp, true);//has ver scrollbar
			
			if (sh > ch + 3) {//3 is border
				if (hsb && !vsb)
					jq(editorcmp).css('height', jq.px0(sh + zss.Spreadsheet.scrollWidth));// extend height
				else
					jq(editorcmp).css('height', jq.px0(sh));
			}
			if (sh > ch + 3 || forceadj) {
				var custColWidth = local.sheet.custColWidth,
					custRowHeight = local.sheet.custRowHeight;
				local.sw = custColWidth.getCellIndex(custColWidth.getStartPixel(local.col) + sw)[0] - local.col;
				local.sh = custRowHeight.getCellIndex(custRowHeight.getStartPixel(local.row) + sh)[0] - local.row;
			}			
		}, 0);
	}
});