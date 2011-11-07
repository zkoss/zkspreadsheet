/* Cell.js

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
	//condition: cell type is string, halign is left, no wrap, no merge
	function evalOverflow(ctrl) {
		return ctrl.cellType == STR_CELL && !ctrl.wrap && !ctrl.merid 
				&& ctrl.halign == 'l'; //&& ctrl.sheet.config.textOverflow
	}

var NUM_CELL = 0,
	STR_CELL = 1,
	FORMULA_CELL = 2,
	BLANK_CELL = 3,
	BOOLEAN_CELL = 4,
	ERROR_CELL = 5,
	Cell = 
zss.Cell =  zk.$extends(zk.Widget, {
	widgetName: 'Cell',
	/**
	 * The data source of the cell
	 */
	src: null,
	/**
	 * Cell text
	 */
	text: '',
	/**
	 * Whether cell is locked or not
	 * 
	 * Default: true
	 */
	lock: true,
	/**
	 * Indicate whether should invoke processd overflow on cell or not
	 * 
	 * Currently, supports left aligned cell only
	 */
	overflow: false,
	/**
	 * Re-process overflow
	 */
	redoOverflow: false,
	/**
	 * Indicate current cell's width is insufficient, overflow cell content to neighbor cells
	 */
	//overflowed,
	/**
	 * Indicate current cell is overlap by a overflowed cell
	 */
	//overlapBy: null,
	/**
	 * Cell type
	 * 
	 * <ul>
	 * 	<li>0: Numeric Cell type</li>
	 * 	<li>1: String Cell type</li>
	 * 	<li>2: Formula Cell type</li>
	 * 	<li>3: Blank Cell type</li>
	 * 	<li>4: Boolean Cell type</li>
	 * 	<li>5: Error Cell type</li>
	 * </ul>
	 * 
	 * Default: Blank Cell type 
	 */
	cellType: 3,
	/**
	 * Horizontal alignment for the cell
	 * 
	 * <ul>
	 * 	<li>l: align left</li>
	 * 	<li>c: align center</li>
	 * 	<li>r: align right</li>
	 * </ul>
	 * 
	 * Default: align left
	 */
	halign: "l",
	/**
	 * Vertical alignment for the cell
	 * 
	 * <ul>
	 * 	<li>t: align top</li>
	 * 	<li>c: align center</li>
	 * 	<li>b: align bottom</li>
	 * </ul>
	 * 
	 * Default: align top
	 */
	valign: 't',
	/**
	 * Whether the text should be wrapped
	 * 
	 * Default: false
	 */
	wrap: false,
	/**
	 * Whether listen onRowHeightChanged event or not 
	 * Currently, use only on IE6/IE7 for vertical align
	 * 
	 * Default: false
	 */
	//_listenRowHeightChanged: false,
	/**
	 * Whether listen sheet onRowAdded event or not 
	 * 
	 * Default: false
	 */
	//_listenRowAdded: false
	$init: function (sheet, block, row, col, src) {
		this.$supers(zss.Cell, '$init', []);
		
		this.sheet = sheet;
		this.block = block;
		this.r = row;
		this.c = col;
		this.src = src;
		var data = src.getRow(row).getCell(col);
		
		this.text = data.text ? data.text : '';
		this.edit = data.editText ? data.editText : '';
		this.hastxt = !!this.text;
		
		this.zsw = data.widthId;
		this.zsh = data.heightId;
		
		this.lock = data.lock;
		this.cellType = data.cellType;
		
		this.halign = data.halign;
		this.valign = data.valign;
		this.rborder = data.rightBorder;

		var mId = data.mergeId;
		if (mId) {
			var r = data.merge;
			this.merid = mId;
			this.mergeId = mId;
			
			this.merl = r.left;
			this.merr = r.right;
			this.mert = r.top;
			this.merb = r.bottom;
			this.mergeCls = data.mergeCls;
		}
		this.wrap = data.wrap;
		this.heightId = data.heightId;
		this.widthId = data.widthId;
		
		this.style = data.style;
		this.innerStyle = data.innerStyle;
	},
	doClick_: function (evt) {
		//do nothing. eat the event.
	},
	doMouseDown_: function (evt) {
		this.sheet._doMousedown(evt);
	},
	doMouseUp_: function (evt) {
		this.sheet._doMouseup(evt);
	},
	doDoubleClick_: function (evt) {
		this.sheet._doMousedblclick(evt);
	},
	/**
	 * Returns whether cell locked or not
	 * @return boolean
	 */
	isLocked: function () {
		return this.lock;
	},
	/**
	 * Update cell
	 */
	update_: function () {
		var r = this.r,
			c = this.c,
			data = this.src.getRow(r).getCell(c) || this.sheet._wgt._activeRange.getRow(r).getCell(c),
			edit = data.editText,
			format = data.formatText,
			lock = data.lock,
			wrap = data.wrap,
			st = this.style = data.style,
			ist = this.innerStyle = data.innerStyle,
			n = this.comp,
			txtNode = this.txtcomp;
		if (n.style.cssText != st || txtNode.style.cssText != ist) {
			n.style.cssText = st;
			txtNode.style.cssText = ist;
		}
		this.cellType = data.cellType;
		this.halign = data.halign;
		this.valign = data.valign;
		this.rborder = data.rightBorder;
		this.edit = data.editText;
		this.redoOverflow = this._updateListenOverflow(evalOverflow(this));
		
		this.setText(data.text);
	},
	/**
	 * Set overflow attribute and register listener or unregister onProcessOverflow listener base on overflow attribute
	 * @param boolean 
	 * @return boolean whether reset overflow attribute or not
	 */
	_updateListenOverflow: function (b) {
		var curr = !!this.overflow;
		if (curr != b) {
			this.sheet[curr ? 'unlisten' : 'listen']({onProcessOverflow: this.proxy(this._onProcessOverflow)});
			this.overflow = b;
			return true;
		}
		return false;
	},
	_updateHasTxt: function (bool) {
		this.hastxt = bool;
		jq(this.comp).css('z-index', bool ? "" : 1);
	},
	/**
	 * Sets the text of the cell
	 * 
	 * @param string text
	 * @param boolean whether fire sheet overflow event or not 
	 */
	setText: function (txt, noTriggerOverflowEvt) {
		if (!txt)
			txt = "";
		var oldTxt = this.getText(),
			difTxt = txt != oldTxt;
		this._updateHasTxt(txt != "");
		this._setText(txt);

		if (this.overflow || this.redoOverflow) {
			if (difTxt) {
				Cell._processOverflow(this);
				if (!noTriggerOverflowEvt)
					this.sheet.triggerOverflowColumn_(this.r, this.c);
				this.redoOverflow = false;
			}
		}
	},
	_updateVerticalAlign: function () {
		var	v = this.valign,
			text = this.text,
			txtcomp = this.txtcomp;
		if (txtcomp.style.display == 'none' || !text)
			return;
		var	$n = jq(txtcomp),
			$txtInner = zk.ie6_ || zk.ie7_ ? $n.children('div') : null;
		if ($txtInner)
			$n.css('position', 'relative');
		else
			return;
		
		switch (v) {
		case 't':
			$txtInner.css({'top': "0px", 'bottom': ''});
			break;
		case 'c':
			var ch = $n.height(),
				ich = $txtInner.height();
			if (!ch || !ich)
				return;
			var	ah = (ch - ich) / 2,
				p = Math.ceil(ah * 100 / ch);
			if (p)
				$txtInner.css({'top': p + "%", 'bottom': ''});
			break;
		case 'b':
			$txtInner.css({'bottom': "0px", 'top': ''});	
			break;
		}
	},
	_onRowHeightChanged: function (evt) {
		if (evt.data.row == this.r)
			this._updateVerticalAlign();
	},
	_updateListenRowHeightChanged: function (b) {
		var curr = !!this._listenRowHeightChanged;
		if (curr != b) {
			this.sheet[curr ? 'unlisten' : 'listen']({onRowHeightChanged: this.proxy(this._onRowHeightChanged)});
			this._listenRowHeightChanged = b;
		}
	},
	getText: function () {
		var n = this.txtcomp;
		n = zk.ie6_ || zk.ie7_ ? n.firstChild : n;
		return n.innerHTML;
	},
	getPureText: function () { //feature #26: Support copy/paste value to local Excel
		var n = this.txtcomp;
		n = zk.ie6_ || zk.ie7_ ? n.firstChild : n;
		return n.textContent || n.innerText;
	},
	_setText: function (txt) {
		if (!txt)
			txt = "";
		var n = this.txtcomp;
		this.text = zk.ie6_ || zk.ie7_ ? n.firstChild.innerHTML = txt : n.innerHTML = txt;
		
		if (zk.ie6_ || zk.ie7_) {
			if (txt) {
				this._updateVerticalAlign();
			}
			this._updateListenRowHeightChanged(!!txt);
		}
	},
	redraw: function (out) {
		out.push(this.getHtml());
	},
	getHtml: function () {
		var	uid = this.uuid,
			text = this.text,
			style = this.style,
			innerStyle = this.innerStyle;
		return '<div id="' + uid + '" class="' + this.getZclass() + '" zs.t="SCell" '
				+ (style ? 'style="' +  style + '"' : '') + '><div id="' + uid + '-real" class="' +
				this._getInnerClass() + '" ' + (innerStyle ? 'style="' + innerStyle + '"' : '') + 
				'>' + (zk.ie6_ || zk.ie7_ ? '<div style="left:0px;position:absolute;width:100%">' + text + '</div>' : text) + '</div></div>';
	},
	_onRowAdded: function () {
		if (this.overflow) {
			Cell._processOverflow(this);
		}
		if (this.hastxt && (zk.ie6_ || zk.ie7_) && this.valign != 't') {
			this._updateVerticalAlign();
		}
	},
	bind_: function (desktop, skipper, after) {
		this.$supers(zss.Cell, 'bind_', arguments);
		
		var n = this.comp = this.$n(),
			sheet = this.sheet;
		n.ctrl = this;
		this.txtcomp = n.firstChild;
		
		this._updateListenOverflow(evalOverflow(this));
		var hasText = !!this.text,
			processValign = false;
		this._updateHasTxt(hasText);
		if (hasText && (zk.ie6_ || zk.ie7_) && this.valign != 't') { //ie6 / ie7 doesn't suport css vertical align
			processValign = true;
			this._updateListenRowHeightChanged(true);
		}
		if (this.overflow || processValign) {
			this._listenRowAdded = true;
			this.parent.listen({onRowAdded: this.proxy(this._onRowAdded)});
		}
	},
	unbind_: function () {
		if (this._listenRowAdded)
			this.parent.unlisten({onRowAdded: this.proxy(this._onRowAdded)});

		this._updateListenOverflow(false);
		this._updateListenRowHeightChanged(false);
		
		this._updateHasTxt(false);
		if (this.overflowed) {
			//remove the overlap relation , because the cell that is overlaped maybe not be destroied.
			Cell._clearOverlapRelation(this);
		}
		
		this.comp = this.comp.ctrl = this.txtcomp = this.sheet = this.overlapBy = this._listenRowHeightChanged =
		this.block = this.lock = null;
		
		this.$supers(zss.Cell, 'unbind_', arguments);
	},
	doMouseDown_: function (evt) {
		this.sheet._doMousedown(evt);
	},
	doMouseUp_: function (evt) {
		this.sheet._doMouseup(evt);
	},
	_onProcessOverflow: function (evt) {
		if (this.overflow) {
			var row = this.r,
				data = evt.data;
			if (data) {// for onProcessOverflow event
				var tRow = data.tRow,
					bRow = data.bRow;
				if (this.c < data.col) {
					if ((tRow == undefined && bRow == undefined) || 
						(tRow && bRow && row >= tRow && row <= bRow)) {
						Cell._processOverflow(this);
					}
				}
			} else // onRowAdded event won't have row, col data
				Cell._processOverflow(this);
		}
	},
	//super//
	getZclass: function () {
		var cls = 'zscell',
			hId = this.heightId,
			wId = this.widthId,
			mId = this.mergeId,
			mCls = this.mergeCls;
		if (hId)
			cls += (' zshi' + hId);
		if (wId)
			cls += (' zsw' + wId);
		if (mCls)
			cls += mCls;
		return cls;
	},
	_getInnerClass: function () {
		var cls = 'zscelltxt',
			hId = this.heightId,
			wId = this.widthId;
		if (hId)
			cls += (' zshi' + hId);
		if (wId)
			cls += (' zswi' + wId);
		return cls;
	},
	/**
	 * Sets the width position index
	 * @param int zsw the width position index
	 */
	appendZSW: function (zsw) {
		this.zsw = zsw;
		jq(this.comp).addClass("zsw" + zsw);
		jq(this.txtcomp).addClass("zswi" + zsw);
	},
	/**
	 * Sets the height position index
	 * @param int zsh the height position index
	 */
	appendZSH: function (zsh) {
		this.zsh = zsh;
		jq(this.comp).addClass("zshi" + zsh);
		jq(this.txtcomp).addClass("zshi" + zsh);
	},
	/**
	 * Sets the column index of the cell
	 * @param int column index
	 */
	resetColumnIndex: function (newcol) {
		this.c = newcol;
	},
	/**
	 * Sets the row index of the cell
	 * @param int row index
	 */
	resetRowIndex: function (newrow) {
		this.r = newrow;
	}
}, {
	/**
	 * Update cell's text and style
	 * @param zss.Cell ctrl
	 * @param map parameter setting
	 */
	updateCell: function (ctrl, parm) {
		var st = parm.st,//style
			ist = parm.ist,//style of inner div
			hal = parm.hal,
			vtal = parm.vtal,
			lock = parm.lock,
			cellType = parm.ctype,
			cmp = ctrl.comp;
		ctrl.cellType = cellType != undefined ? cellType : BLANK_CELL;
		ctrl.lock = !!lock;//default is locked
		ctrl.wrap = parm.wrap;
		st = !st ? '' : st;
		ist = !ist ? '' : ist;
		if (cmp.style.cssText != st || ctrl.txtcomp.style.cssText != ist) {
			cmp.style.cssText = st;
			ctrl.txtcomp.style.cssText = ist;
		}

		ctrl.halign = !hal ? "l" : hal;//default horizontal align is left
		ctrl.valign = !vtal ? 't' : vtal;//default vertical align is top
		ctrl.rborder = !!parm.rbo;
		ctrl.edit = parm.edit;
		ctrl._updateListenOverflow(evalOverflow(ctrl));
		ctrl.setText(parm.txt);
	},
	_clearOverlapRelation: function (ctrl) {
		var cmp = ctrl.comp,
			sheet = ctrl.sheet;
		if (this.overflowed && !sheet.invalid) {
			var next = cmp.nextSibling,
				nCtrl;
			while (next) {
				nCtrl = next.ctrl;
				if (nCtrl && nCtrl.overlapBy == ctrl) {
					nCtrl.overlapBy = null;
					next = next.nextSibling;
					continue;	
				}
				break;
			}
		}
	},
	/**
	 * Invoke after CSS ready, set alignment
	 */
	_processVerticalAlign: function (ctrl) {
		if (ctrl.invalid)
			return;
		
		ctrl._updateVerticalAlign();
	},
	_processOverflow: function (ctrl) {		
		var sheet = ctrl.sheet;
		if (ctrl.invalid || !sheet.config.textOverflow) {
			//a cell might load on demand, it will cause a initial later to process text overflow(see also,CellBlock._createCell)
			//but it some time, it will be prune after load ,
			//so we must ignore this when ctrl.invliad
			return;
		}
		var cmp = ctrl.comp,
			txtcmp = ctrl.txtcomp;
		jq(txtcmp).css({'width': '', 'position': ''});//remove old value.
		var sw = txtcmp.scrollWidth,
			cellPad = sheet.cellPad,
			oldOverlapBy = ctrl.overlapBy,
			overflow = ctrl.overflow && sw > (cmp.clientWidth - 2 * cellPad);
		if (overflow) {
			ctrl.overflowed = true;
			txtcmp.style.position = zk.ie || zk.safari ? "absolute" : "relative";
			var w = cmp.clientWidth,
				prev = cmp,
				next = cmp.nextSibling;
			while (next) {
				var nCtrl = next.ctrl,
					rBorder = prev.ctrl.rborder;

				nCtrl.overlapBy = null;
				if (nCtrl.hastxt) {
					jq(prev).removeClass(rBorder ? "zscell-over-b" : "zscell-over");
					break;
				}
				
				jq(prev).removeClass(rBorder ? "zscell-over" : "zscell-over-b")
						.addClass(rBorder ? "zscell-over-b" : "zscell-over");
				
				var ow = next.offsetWidth;
				w = w + ow;
				if (w > sw) {
					jq(next).removeClass(next.ctrl.rborder ? "zscell-over-b" : "zscell-over");
					
					//keep remove next overflow if it is over by this cell
					next = next.nextSibling;
					while (next) {
						if (ctrl != next.ctrl.overlapBy)
							break;
						jq(next).removeClass(next.ctrl.rborder ? "zscell-over-b" : "zscell-over");

						next.ctrl.overlapBy = null;
						next = next.nextSibling;
					}
					break;
				}
				prev = next;
				next = next.nextSibling;//jq(next).next('DIV')[0];
			}
			w =  w - cellPad;
			jq(txtcmp).css('width', jq.px0(w));//when some column or row is invisible, the w might be samlle then zero
		} else {
			ctrl.overflowed = false;
			jq(cmp).removeClass(cmp.ctrl.rborder ? "zscell-over-b" : "zscell-over");
			
			var next = cmp.nextSibling;
			while (next) {
				//next no initial yet skip to check
				if(!next.ctrl || 
					ctrl != next.ctrl.overlapBy && (oldOverlapBy ? oldOverlapBy != next.ctrl.overlapBy : true))
					break;
				jq(next).removeClass(next.ctrl.rborder ? "zscell-over-b" : "zscell-over");

				next.ctrl.overlapBy = null;
				next = next.nextSibling;
			}
		}
		if (oldOverlapBy)
			Cell._processOverflow(oldOverlapBy);
	}
});
})();