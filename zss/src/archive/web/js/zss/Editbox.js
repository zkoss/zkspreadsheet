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

(function () {

	function newLine(n) {
		var sel = zk(n).getSelectionRange();	
		if (sel[0] > sel[1]) {
			var t = sel[1];
			sel[1] = sel[0];
			sel[0] = t;
		}
		var str = n.value,
			len = str.length,
			str1 = str.substring(0, sel[0]); 
		str = str1 +'\n'+ str.substring(sel[1], len);
		n.value = str;
		
		var pos = sel[0];
		if (n.setSelectionRange) { 
			pos = zk.opera ? pos + 2 : pos + 1;
			n.setSelectionRange(pos, pos);
		} else if (n.createTextRange) {//IE
			var range = n.createTextRange(),
				i = -1;
			pos = pos + 1;
			while((i = str1.indexOf('\n',i+1)) >= 0) {
				pos--;
			}
			range.move('character', pos);
			range.select();
		}
	}
	
	function syncEditorHeight(n) {
		var ch = n.clientHeight,
			cw = n.clientWidth,
			sw = n.scrollWidth,
			sh = n.scrollHeight,
			hsb = zkS._hasScrollBar(n),//has hor scrollbar
			vsb = zkS._hasScrollBar(n, true);//has ver scrollbar
		
		if (sh > ch + 3) {//3 is border
			if (hsb && !vsb)
				jq(n).css('height', jq.px0(sh + zss.Spreadsheet.scrollWidth));// extend height
			else
				jq(n).css('height', jq.px0(sh));
		}
	}

	/**
	 * Returns cell reference
	 * 
	 * @param string src the source to extract reference
	 * @return array cell reference
	 */
	function getCellRefs(src) {
		return src.match('\\$?([A-Za-z]+)\\$?([0-9]+)');
	}
	
	/**
	 * Returns the editing formula info
	 *  
	 * @param string type the type of editor (optional)
	 * @param Object n the editor DOM element
	 * @param array p the text cursor position (optional)
	 */
	function getEditingFormulaInfo(type, n, p) {
		var p = p || zk(n).getSelectionRange(),
			tp = type || 'inlineEditing',
			start = p[0],
			end = p[1],
			v = n.value;
		if (start != end) { //text has selection, need to replace cell reference
			return {start: start, end: end, type: tp};
		} else {
			if (!start) {
				return {start: 0, end: 0, type: tp};
			} else {
				var DELIMITERS = ['=', '+', '-', '*', '/', '!', ':', '^', '&', '(',  ',', '.'],
					i = start - 1,
					c = v.charAt(i);
				if (DELIMITERS.$contains(c)) {
					return {start: start, end: start, type: tp};
				}
			}
		}
	}
	
	/**
	 * Insert cell reference to editor
	 * 
	 * @param Object sheet the ({@link zss.SSheetCtrl})
	 * @param Object n the editor DOM element
	 * @param int tRow
	 * @param int lCol
	 * @param int bRow
	 * @param int rCol
	 */
	function insertCellRef(sheet, n, tRow, lCol, bRow, rCol) {
   		var info = sheet.editingFormulaInfo,
			start = info.start,
			end = info.end,
			v = n.value,
			c1 = sheet.getCell(tRow, lCol),
			c2 = tRow == bRow && lCol == rCol ? null : sheet.getCell(bRow, rCol),
			ref = c1.ref;
		if (c2) {
			ref += (':' + c2.ref);
		}
		if (!start) {
			ref = '=' + ref; 
		}
		n.value = v.substring(0, start) + ref + v.substring(end, v.length);
		sheet.inlineEditor.setValue(n.value);
		end = start + ref.length;
		info.end = end;
		if (zk.ie) { 
			setTimeout(function () {
				zk(n).setSelectionRange(end, end);
			});
		} else {
			zk(n).setSelectionRange(end, end);
		}
	}

zss.FormulabarEditor = zk.$extends(zul.inp.InputWidget, {
	widgetName: 'FormulabarEditor',
	row: -1,
	col: -1,
   	$init: function (wgt) {
   		this.$supers(zss.FormulabarEditor, '$init', []);
   		this._wgt = wgt;
   	},
   	bind_: function () {
   		this.$supers(zss.FormulabarEditor, 'bind_', arguments);
   		var sheet = this._wgt.sheetCtrl;
   		if (sheet) {
   			this.sheet = sheet;
   			sheet.formulabarEditor = this;
   			
   			sheet.listen({'onStartEditing': this.proxy(this._onStartEditing)})
   				.listen({'onStopEditing': this.proxy(this.stop)})
   				.listen({'onCellSelection': this.proxy(this._onCellSelection)})
   				.listen({'onContentsChanged': this.proxy(this._onContentsChanged)});
   		}
   	},
   	unbind_: function () {
   		var sheet = this.sheet;
   		if (sheet) {
   			
   			sheet.unlisten({'onStartEditing': this.proxy(this._onStartEditing)})
   				.unlisten({'onStopEditing': this.proxy(this.stop)})
   				.unlisten({'onCellSelection': this.proxy(this._onCellSelection)})
   				.unlisten({'onContentsChanged': this.proxy(this._onContentsChanged)});
   			
   			sheet.formulaEditor = null;
   		}
   		this.sheet = this._wgt = null;
   		this.$supers(zss.FormulabarEditor, 'unbind_', arguments);
   	},
   	disable: zk.$void,
   	cancel: function () {
   		var sheet = this.sheet;
   		if (sheet) {
   			sheet.inlineEditor.cancel();
   		}
   		this.row = this.col = -1;
   	},
   	stop: function () {
	   	var sheet = this.sheet;
	   	if (sheet) {
	   		sheet.inlineEditor.stop();
   		}
   		this.row = this.col = -1;
   	},
   	/**
   	 * Sync editor top position base on text cursor position
   	 * 
   	 * @param int p the text cursor position. If doesn't specify, use current text cursor position
   	 */
   	syncEditorPosition: function (p) {
		var n = this.$n('real');
		if (p == undefined) {
			p = zk(n).getSelectionRange()[1];
		}
		if (!p) {
			jq(n).css('top', '0px');
			return;
		}
		var	lineHeight = zk.parseInt(jq(n).css('line-height')),
			str = n.value,
			len = str.length,
			lineNum = 1;
		for (var i = 0; i < len; i++) {
			if (i == p)
				break;
			if (str.charAt(i) == '\n')
				lineNum++;
		}
		if (lineNum - 1) {
			var h = lineNum * lineHeight,
				curHgh = jq(this.$n()).height(),
				curLine = Math.floor(curHgh / lineHeight);
			if (curLine < lineNum) {
				//6: n.paddingTop (2px) + n.marginTop (1px) + text top offset (3px)
				jq(n).css('top', (curHgh - h - 6) + 'px');
			}
		} else {
			jq(n).css('top', '0px');
		}
   	},
   	/**
   	 * Sets the value of formula editor
   	 * 
   	 * @string string value
   	 * @param int pos the text cursor position
   	 */
   	setValue: function (v, pos) {
	   	var n = this.$n('real');
   	   	n.value = v;
   		if (this.isRealVisible()) {
   	   		syncEditorHeight(n);
   	   		
   	   		var p = pos || n.value.length;
   	   		this.syncEditorPosition(p);
   		}
   	},
   	getValue: function () {
   		return this.$n('real').value;
   	},
   	newLine: function () {
   		newLine(this.$n('real'));
   	},
   	/**
   	 * Focus on editor
   	 */
   	focus: function () {
   		var n = this.$n('real');
   		if (zk(n).isRealVisible(true))
   			n.focus();
   	},
   	doFocus_: function () {
   		var sheet = this.sheet;
   		if (sheet) {
   			var info = sheet.editingFormulaInfo;
   			if (info) {
   				if (!info.moveCell) {//focus back by user's mouse evt, re-eval editing formula info
   					sheet.editingFormulaInfo = getEditingFormulaInfo('formulabarEditing', this.$n('real'));
   				} else {
   					//focus back to editor while move cell, do nothing
   				}
   			} else {
   	   	   		var p = this.sheet.getLastFocus(),
				 	ls = sheet.selArea.lastRange;
				sheet.moveCellFocus(p.row, p.column);
				sheet.moveCellSelection(ls.left, ls.top, ls.right, ls.bottom, false, true);
   			}
   		}
   	},
   	doBlur_: function () {
	   	var sheet = this.sheet;
	   	if (sheet) {
	   		setTimeout(function () {
	   			if (!sheet._wgt.hasFocus()) {
	   				sheet.dp.stopEditing(sheet.innerClicking > 0 ? "refocus" : "lostfocus");
	   			}
	   		});
	   	}
   	},
   	_onContentsChanged: function (evt) {
   		var sheet = this.sheet;
   		if (sheet && this.isRealVisible()) {
   			var p = sheet.getLastFocus(),
   				c = sheet.getCell(p.row, p.column);
   			if (c) {
   				this.$n('real').value = c.edit || '';
   			}
   		}
   	},
   	_onStartEditing: function (evt) {
   		var d = evt.data;
   		this.row = d.row;
   		this.col = d.col;
   	},
   	_onCellSelection: function (evt) {
   		var sheet = this.sheet;
   		if (sheet) {
   			if (sheet.state == zss.SSheetCtrl.FOCUSED) {
   			   	var d = evt.data,
   			   		cell = sheet.getCell(d.top, d.left);
   			   	if (cell) {
   			   		this.$n('real').value = cell.edit || '';
   			   		this.syncEditorPosition(0);
   			   	} else {
   			   	}
   			} else if (sheet.state == zss.SSheetCtrl.EDITING) {
   				var info = sheet.editingFormulaInfo;
   				if (info && 'formulabarEditing' == info.type) {
   					var d = evt.data;
   					insertCellRef(sheet, this.$n('real'), d.top, d.left, d.bottom, d.right);
   				}
   			}
   		}
   	},
	doMouseDown_: function (evt) {
		var sheet = this.sheet;
		if (sheet) {
			if (sheet.state == zss.SSheetCtrl.FOCUSED) {
				var p = sheet.getLastFocus(),
					c = sheet.getCell(p.row, p.column);
				if (c && sheet.dp.startEditing(evt, c.edit, 'formulabarEditing')) {
					this.row = p.row;
					this.col = p.column;
				} else { //not allow edit, focus back
					sheet.dp.gainFocus(true);
				}
			} else if (sheet.state == zss.SSheetCtrl.NOFOCUS) {
				var p = sheet.getLastFocus(),
					row = p.row,
					col = p.column;
				sheet.dp.moveFocus(row == -1 ? 0 : row, col == -1 ? 0 : col);
				evt.stop();
			} else if (sheet.state == zss.SSheetCtrl.EDITING) {
				var info = sheet.editingFormulaInfo;
				if (info && info.moveCell) {//re-eval editing formula info
					sheet.editingFormulaInfo = getEditingFormulaInfo('formulabarEditing', this.$n('real'));
				}
			}
		}
	},
	doKeyDown_: function (evt) {
		var sheet = this.sheet;
		if (sheet) {
			sheet._doKeydown(evt);
		}
	},
	afterKeyDown_: function (evt, simulated) {//must eat the event, otherwise cause delete key evt doesn't work correctly
	},
	doKeyUp_: function (evt) {
		var sheet = this.sheet,
			keycode = evt.keyCode;
		if (sheet && sheet.state == zss.SSheetCtrl.EDITING) {
			var n = this.$n('real'),
				p = zk(n).getSelectionRange(),
				end = p[1],
				v = n.value,
				info = sheet.editingFormulaInfo;
			syncEditorHeight(n);
			this.syncEditorPosition(end);
			sheet.inlineEditor.setValue(v);
			this._wgt.fire('onEditboxEditing', {token: '', sheetId: sheet.serverSheetId, clienttxt: n.value});
			
			if (!v) {
				sheet.editingFormulaInfo = null;
				return;
			}
			
			var tp = 'formulabarEditing';
			if (info) {
				info.type = tp;
				if (info.moveCell) {
					if (info.end != end) { //means keyboard input new charcater
						sheet.editingFormulaInfo = getEditingFormulaInfo(tp, n, p);
					} else {
						//no new charcater, remian the same editing info, do nothing.
					}
				} else {
					sheet.editingFormulaInfo = getEditingFormulaInfo(tp, n, p);
				}
			} else {
				sheet.editingFormulaInfo = getEditingFormulaInfo(tp, n, p);
			}
		}
	},
	doKeyPress_: function (evt) {//eat the event
	},
   	setWidth: function (v) {
   		this.$supers(zss.FormulabarEditor, 'setWidth', arguments);
   		var w = this.$n().clientWidth;
   		this.$n('real').style.width = jq.px(zk.ie ? w - 8: w); //8: IE's textarea scrollbar width
   	},
   	setHeight: function (v) {
   		this.$supers(zss.FormulabarEditor, 'setHeight', arguments);
   		var h = this.$n().clientHeight,
   			n = this.$n('real');
   		n.style.height = n.scrollHeight > h ? n.scrollHeight + 'px' : h + 'px';
   		this.syncEditorPosition();
   	},
   	redrawHTML_: function () {
   		var uid = this.uuid,
   			zcls = this.getZclass();
   		return '<div id="' + uid + '" class="' + zcls + '"><textarea id="' + uid + '-real" class="' + zcls + '-real">' + this._areaText() + '</textarea></div>';
   	},
   	getZclass: function () {
   		return 'zsformulabar-editor';
   	}
});

/**
 * Editbox represent a edit area for cell
 */
zss.Editbox = zk.$extends(zul.inp.InputWidget, {
	widgetName: 'InlineEditor',
	_editing: false,
	_type: 'inlineEditing',
	$init: function (sheet) {
		this.$supers(zss.Editbox, '$init', []);
		var wgt = this._wgt = sheet._wgt;
		
		this.sheet = sheet;
		this.row = -1;
		this.col = -1;
		this.disabled = true;
	},
	bind_: function () {
		this.$supers(zss.Editbox, 'bind_', arguments);
		this.comp = this.$n();
		this.comp.ctrl = this;
		
		this.sheet.listen({'onStartEditing': this.proxy(this._onStartEditing)})
   			.listen({'onStopEditing': this.proxy(this.stop)})
   			.listen({'onCellSelection': this.proxy(this._onCellSelection)});
	},
	unbind_: function () {
		this.sheet.unlisten({'onStartEditing': this.proxy(this._onStartEditing)})
   			.unlisten({'onStopEditing': this.proxy(this.stop)})
   			.unlisten({'onCellSelection': this.proxy(this._onCellSelection)});
		
		this.sheet = this.comp.ctrl = this.comp = this._wgt = null;
		this.$supers(zss.Editbox, 'unbind_', arguments);
	},
	focus: function () {
		var n = this.comp;
   		if (zk(n).isRealVisible(true))
   			n.focus();
	},
	isEditing: function () {
		return this._editing;
	},
   	_onStartEditing: function (evt) {
   		var d = evt.data;
   		this.row = d.row;
   		this.col = d.col;
   	},
	_onCellSelection: function (evt) {
   		var sheet = this.sheet;
   		if (sheet) {
   			if (sheet.state == zss.SSheetCtrl.EDITING) {
   				var info = sheet.editingFormulaInfo;
   				if (info && 'inlineEditing' == info.type) {
   					var d = evt.data,
   						formulabarEditor = sheet.formulabarEditor;
   					insertCellRef(sheet, this.comp, d.top, d.left, d.bottom, d.right);
   					if (formulabarEditor) {
   						var n = this.comp,
   							v = n.value;
   						formulabarEditor.setValue(v);
   					}
   				}
   			}
   		}
	},
	doFocus_: function () {
		this._editing = true;
	},
	doBlur_: function () {
	   	var sheet = this.sheet;
	   	if (sheet) {
	   		setTimeout(function () {
	   			if (!sheet._wgt.hasFocus()) {
	   				sheet.dp.stopEditing(sheet.innerClicking > 0 ? "refocus" : "lostfocus");
	   			}
	   		});
	   	}
	},
	doMouseDown_: function (evt) {
		var sheet = this.sheet;
		if (sheet) {
			if (sheet.state == zss.SSheetCtrl.EDITING) {
				var info = sheet.editingFormulaInfo;
				if (info && info.moveCell) {//re-eval editing formula info
					sheet.editingFormulaInfo = getEditingFormulaInfo(this._type, this.$n());
				}
			}
		}
	},
	doKeyPress_: function (evt) {
		this.autoAdjust(true);
	},
	doKeyDown_: function (evt) {
		if (this.disabled)
			evt.stop();
		else {
			this.sheet._doKeydown(evt);
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
		}
	},
	afterKeyDown_: function (evt, simulated) {//must eat the event, otherwise cause delete key evt doesn't work correctly
	},
	doKeyUp_: function (evt) {
		var sheet = this.sheet;
		if (sheet && sheet.state == zss.SSheetCtrl.EDITING) {
			var	sheet = this.sheet,
				formulabarEditor = sheet.formulabarEditor,
				value = sheet.inlineEditor.getValue(),
				keycode = evt.keyCode,
				info = sheet.editingFormulaInfo,
				n = this.comp,
				p = zk(n).getSelectionRange(),
				end = p[1];
			if (formulabarEditor) {
				formulabarEditor.setValue(value, end);
			}
			sheet._wgt.fire('onEditboxEditing', {token: '', sheetId: sheet.serverSheetId, clienttxt: value});
			
			if (!value) {
				sheet.editingFormulaInfo = null;
				return;
			}
			
			var tp = this._type;
			if (info) {
				info.type = tp;
				if (info.moveCell) {
					if (info.end != end) { //means keyboard input new charcater
						sheet.editingFormulaInfo = getEditingFormulaInfo(tp, n, p);
					} else {
						//no new charcater, remian the same editing info, do nothing.
					}
				} else {
					sheet.editingFormulaInfo = getEditingFormulaInfo(tp, n, p);
				}
			} else {
				sheet.editingFormulaInfo = getEditingFormulaInfo(tp, n, p);
			}
		}
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
	 * Sets the value of inline editor
	 * 
	 * @param string value
	 */
	setValue: function (v) {
		if (this.disabled) {
			return;
		}
		this.comp.value = v;
		this.autoAdjust(true);
	},
	_startEditing: function (noFocus) {
		var sheet = this.sheet;
		if (sheet) {
			this._editing = !noFocus;
			var formulabarEditor = sheet.formulabarEditor,
				n = this.comp,
				v = n.value;
			if (formulabarEditor) {
				formulabarEditor.setValue(v);
			}
			if ('=' == v.charAt(0)) {
				var p = v.length;
				sheet.editingFormulaInfo = getEditingFormulaInfo(this._type, n, [p, p]);	
			}
		}
	},
	/**
	 * 
	 * @cellcmp
	 * @row int row number (0 based)
	 * @col int column number (0 based)
	 * @param string value the edit value 
	 * @param boolean noFocus whether focus on inline editor or not
	 */
	edit: function (cellcmp, row, col, value, noFocus) {
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

		this._startEditing(noFocus);
   			
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
			if (zk.ie) {
				$edit.show();
			}
			if (!noFocus) {
				$edit.focus();
				//issue 230: IE cursor position is not at the text end when press F2
				if (zk.safari || zk.ie) fun();
			}
		}, 25);	
		this.autoAdjust(true);
	},
	cancel: function () {
		this._editing = false;
		this.disable(true);
		jq(this.comp).val('').css('display', 'none');
		this.row = this.col = -1;
	},
	stop: function () {
		this._editing = false;
		this.disable(true);
		var editorcmp = this.comp,
			str = editorcmp.value;
		jq(editorcmp).val('').css('display', 'none');
		this.row = this.col = -1;
		return str;
	},
	getValue: function () {
		return this.comp.value;
	},
	newLine: function () {
		newLine(this.comp);
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
	},
	redrawHTML_: function () {
		return '<textarea id="' + this.uuid + '" class="zsedit" zs.t="SEditbox"></textarea>';
	}
});
})();