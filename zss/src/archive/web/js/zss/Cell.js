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
				&& ctrl.halign == 'l' && ctrl.sheet.config.textOverflow;
	}
	
	//cache
	var NUM_CELL = 0,
		STR_CELL = 1,
		FORMULA_CELL = 2,
		BLANK_CELL = 3,
		BOOLEAN_CELL = 4,
		ERROR_CELL = 5,
		Cell = 
zss.Cell = zk.$extends(zk.Object, {
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
	//valign: 't',
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
	$init: function (sheet, block, cmp, edit) {
		this.$supers('$init', arguments);
		this.id = cmp.id;
		this.sheetid = sheet.sheetid;
		this.sheet = sheet;
		this.block = block;
		this.comp = cmp;
		this.txtcomp = cmp.firstChild;
		this.edit = edit;

		var $n = jq(cmp);
		this.r = zk.parseInt($n.attr('z.r'));
		this.c = zk.parseInt($n.attr('z.c'));

		this.zsw = $n.attr('z.zsw');
		if (this.zsw)
			this.zsw = zk.parseInt(this.zsw);

		this.zsh = $n.attr('z.zsh');
		if (this.zsh)
			this.zsh = zk.parseInt(this.zsh);

		var lock = $n.attr('z.lock'),
			cellType = $n.attr('z.ctype'),
			wrap = this.wrap = $n.attr('z.wrap') == "t",
			halign = $n.attr('z.hal'),
			valign = $n.attr('z.vtal');
		if (lock != undefined)
			this.lock = !(lock == "f");
		if (cellType != undefined)
			this.cellType = cellType;
		if (halign)
			this.halign = halign;
		this.valign = valign ? valign : 't';

		var txt = this.text = 
			zk.ie6_ || zk.ie7_ ? this.txtcomp.firstChild.innerHTML : this.txtcomp.innerHTML;
		//is this cell has right border,
		//i must do some special processing when textoverfow when with rborder.
		this.rborder = $n.attr('z.rbo') == "t";

		//process merge.
		var mid = $n.attr('z.merid');
		if (mid) {
			this.merr = zk.parseInt($n.attr('z.merr')); 
			this.merid = zk.parseInt(mid); //merge id start from 1
			this.merl = zk.parseInt($n.attr('z.merl')); 
			this.mert = zk.parseInt($n.attr('z.mert')); 
			this.merb = zk.parseInt($n.attr('z.merb')); 
		}
		cmp.ctrl = this;
		
		this._updateHasTxt(!!txt);
		if (txt) {
			this._updateHasTxt(true);
			if (zk.ie6_ || zk.ie7_) {
				sheet._wgt._listenRowHeightChanged(this, this._onRowHeightChanged);
				this._listenRowHeightChanged = true;
				sheet.insertSSInitLater(Cell._processVerticalAlign, this);
			}
			if (this.overflow = evalOverflow(this))
				sheet.addSSInitLater(Cell._processOverflow, this);
		}
	},
	/**
	 * Returns whether cell locked or not
	 * @return boolean
	 */
	isLocked: function () {
		return this.lock;
	},
	_updateHasTxt: function (bool) {
		var block = this.block;
		if (this.hastxt && !bool) {
			block._textedcell.$remove(this);	
		} else if (bool && !this.hastxt) {
			block._textedcell.push(this);
		}
		this.hastxt = bool;
		jq(this.comp).css('z-index', bool ? "" : 1);
	},
	/**
	 * Sets the text of the cell
	 * @param string text
	 */
	setText: function (txt) {
		if (!txt)
			txt = "";
		var oldTxt = this.getText(),
			difTxt = txt != oldTxt;
		this._updateHasTxt(txt != "");
		this._setText(txt);

		if (this.overflow) {
			if (this.redoOverflow || difTxt) {
				Cell._processOverflow(this);
				this.redoOverflow = false;
			}
		}
	},
	_setVerticalAlign: function () {
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
			this._setVerticalAlign();
	},
	getText: function () {
		if (this._inner)
			return this._inner.getText();
		return this.txtcomp.innerHTML;
	},
	getPureText: function () { //feature #26: Support copy/paste value to local Excel
		if (this._inner)
			return this._inner.getPureText();
		return this.txtcomp.textContent || this.txtcomp.innerText;
	},
	_setText: function (txt) {
		if (!txt)
			txt = "";

		if (zk.ie6_ || zk.ie7_)
			this.text = this.txtcomp.firstChild.innerHTML = txt;
		else
			this.text = this.txtcomp.innerHTML = txt;
		
		if (zk.ie6_ || zk.ie7_) {
			if (txt) {
				this._setVerticalAlign();
				if (!this._listenRowHeightChanged) {
					this._listenRowHeightChanged = true;
					this.sheet._wgt._listenRowHeightChanged(this, this._onRowHeightChanged);
				}
			} else {
				if (this._listenRowHeightChanged) {
					this.sheet._wgt._unlistenRowHeightChanged(this, this._onRowHeightChanged);
					this._listenRowHeightChanged = false;
				}
			}
		}
	},
	cleanup: function () {
		this.invalid = true;
		
		if (this._listenRowHeightChanged) {
			this.sheet._wgt._unlistenRowHeightChanged(this, this._onRowHeightChanged);
		}
		
		this._updateHasTxt(false);
		if (this.overflowed && !this.block.invalid) {
			//remove the overlap relation , because the cell that is overlaped maybe not be destroied.
			Cell._clearOverlapRelation(this);
		}
		var inner = this._inner;
		if (inner) {
			inner.unbind();
			this._inner = null;
		}
		if (this.comp) this.comp.ctrl = null;
		this.comp = this.txtcomp = this.sheet = this.overlapBy = this._listenRowHeightChanged =
			this.block = this._inner = this._vflex = this.lock = null;
	},
	/**
	 * Sets the width position index
	 * @param int zsw the width position index
	 */
	appendZSW: function (zsw) {
		this.zsw = zsw;
		jq(this.comp).attr("z.zsw", zsw).addClass("zsw" + zsw);
		jq(this.txtcomp).addClass("zswi" + zsw);
	},
	/**
	 * Sets the height position index
	 * @param int zsh the height position index
	 */
	appendZSH: function (zsh) {
		this.zsh = zsh;
		jq(this.comp).attr("z.zsh", zsh).addClass("zshi" + zsh);
		jq(this.txtcomp).addClass("zshi" + zsh);
	},
	/**
	 * Sets the column index of the cell
	 * @param int column index
	 */
	resetColumnIndex: function (newcol) {
		jq(this.comp).attr("z.c", (this.c = newcol));
	},
	/**
	 * Sets the row index of the cell
	 * @param int row index
	 */
	resetRowIndex: function (newrow) {
		jq(this.comp).attr("z.r", (this.r = newrow));
	}
}, {
	/**
	 * Create DOM node and Cell object
	 * @return object Cell
	 */
	createComp: function (sheet, block, parm) {
		var row = parm.row,
			col = parm.col,
			txt = parm.txt,
			st = parm.st,//style
			ist = parm.ist,//style of inner div
			hal = parm.hal,
			vtal = parm.vtal, //vertical align
			merr = parm.merr,
			merid = parm.merid,
			merl = parm.merl,
			mert = parm.mert,
			merb = parm.merb,
			zsw = parm.zsw,
			zsh = parm.zsh,
			cmp = document.createElement("div"),
			drh = parm.drh,
			lock = parm.lock,
			cellType = parm.ctype,
			$n = jq(cmp);

		$n.attr({"zs.t": "SCell", "z.c": col, "z.r": row, 
			"z.hal": (hal ? hal : "l"), "z.vtal": (vtal ? vtal : "t"), 'z.drh': !drh ? '0' : drh});
		
		if (lock != undefined && !lock) $n.attr("z.lock", "f");
		if (cellType != undefined) $n.attr("z.ctype", cellType);
		if (zsw) $n.attr("z.zsw", zsw);
		if (zsh) $n.attr("z.zsh", zsh);
		if (parm.wrap) $n.attr("z.wrap", "t");
		if (parm.rbo) $n.attr("z.rbo", "t");
		if (merid)
			$n.attr({"z.merr": merr, "z.merid": merid, "z.merl": merl, "z.mert": mert, "z.merb": merb}).addClass(merl == col && mert == row? "zsmerge" + merid : mert == row ? "zsmergee" : "zsmergeeu");
		
		if (st)
			cmp.style.cssText = st;
		
		var $txt = jq('<div>' + (zk.ie6_ || zk.ie7_ ? '<div style="left:0px;position:absolute;width:100%"></div>' : '') +'</div>'),
			txtcmp = $txt[0],
			txtInner = zk.ie6_ || zk.ie7_ ? txtcmp.firstChild : null;
		if (ist)
			txtcmp.style.cssText = ist;
		txt = txt ? txt : "";
		txtInner ? txtInner.innerHTML = txt : txtcmp.innerHTML = txt;
		cmp.appendChild(txtcmp);
		
		var sclazz = "zscell" + (zsw ? " zsw" + zsw : "") + (zsh ? " zshi" + zsh : "");
		$n.addClass(sclazz);
		
		sclazz = "zscelltxt" + (zsw ? " zswi" + zsw : "") + (zsh ? " zshi" + zsh : "");
		jq(txtcmp).addClass(sclazz);
		return new zss.Cell(sheet, block, cmp, parm.edit);
	},
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
			ctrl.redoOverflow = true;
		}
		ctrl.halign = !hal ? "l" : hal;//default horizontal align is left
		ctrl.valign = !vtal ? 't' : vtal;//default vertical align is top
		ctrl.rborder = !!parm.rbo;
		ctrl.edit = parm.edit;
		ctrl.overflow = evalOverflow(ctrl);
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
		
		ctrl._setVerticalAlign();
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
			block = ctrl.block,
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
				if (!next.ctrl) {//next no initial yet, then delay and process again.
					zkS.addDelayBatch(zss.Cell._processOverflow, false, ctrl);
					return;
				}
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
			zss.Cell._processOverflow(oldOverlapBy);
	}
});
})();