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
	var Div = zul.wgt.Div;
	zss.BaseInner = zk.$extends(zul.wgt.Div, {
		setSheet: function (sheet) {
			this._sheet = sheet;
		},
		doClick_: function (evt) {
			this._sheet._doMouseleftclick(evt);
		},
		doMouseDown_: function (evt) {
			this._sheet._doMousedown(evt);
		},
		doMouseUp_: function (evt) {
			this._sheet._doMouseup(evt);
		},
		doRightClick_: function (evt) {
			this._sheet._doMouserightclick(evt);
		},
		doDoubleClick_: function (evt) {
			this._sheet._doMousedblclick(evt);
		},
		doKeyDown_: function (evt) {
			this._sheet._doKeyDown(evt);
		},
		afterKeyDown_: function (evt) {
			this._sheet._wgt.afterKeyDown_(evt);
		},
		doKeyPress_: function (evt) {
			this._sheet._doKeypress(evt);
		}
	});

	/**
	 * Cell inner container base on ZK flex for vertical alignment.
	 */
	zss.FlexInner = zk.$extends(zss.BaseInner, {
		$init: function () {
			this.$supers('$init', arguments);
			this.appendChild(new Div());
			this.appendChild(new Div());
			this.appendChild(new Div());
			this.setWidth('100%');
			this.setHeight('100%');
		},
		/**
		 * Sets the text
		 * @param string indicate vertical alignment
		 * @param string indicate horizontal alignment
		 * @param boolean wrap text or not
		 * @param boolean whether the height of the cell is default row height or not
		 */
		setText: function (t, valign, halign, defaultRowHgh) {
			var $t = this.firstChild,
				$c = $t.nextSibling,
				$b = this.lastChild;
			this.valign = valign;
			switch (valign) {
			// set top div visible; hide other div
			case 't':
				$t.$n().innerHTML = t;
				$c.setVisible(false);
				$b.setVisible(false);
				if (!defaultRowHgh)
					this._initShowAndFlex([$t]);
				if (zk.ie6_)
					this._restoreFontSize($t.$n());
				break;
			//set top div flex=true, center div flex=min, bottom div flex=true
			case 'c':
				$c.$n().innerHTML = t;
				$c.setVisible(true);
				if (!defaultRowHgh)
					this._initShowAndFlex([$t, $b], true);
				if (zk.ie6_)
					this._restoreFontSize($c.$n());
				break;
			//set top flex=true; center div invisible, bottom div flex=min
			case 'b':
				$c.setVisible(false);
				$b.$n().innerHTML = t;
				$b.setVisible(true);
				$b.setVflex(-65500);
				if (!defaultRowHgh)
					this._initShowAndFlex([$t], true);
				if (zk.ie6_)
					this._restoreFontSize($b.$n());
			}
			this.fixFlexHeight();
		},
		/**
		 * Fix flexed height
		 */
		fixFlexHeight: zk.ie6_ ? function () {
			var n = this.$n();
			if (!n)
				this._node = n = jq(this.uuid, zk)[0];
			if (n) {
				var cellHgh = n.parentNode.offsetHeight,
					curHgh = n.offsetHeight,
					$t = this.firstChild,
					$c = $t.nextSibling,
					$b = this.lastChild;
				if (curHgh > cellHgh) {
					var differHgh = curHgh - cellHgh;
					switch (this.valign) {
					case 'c':
						var tn = $t.$n(),
							bn = $b.$n(),
							setHgh = Math.round(((tn.offsetHeight + bn.offsetHeight) - differHgh) / 2),
							cssProp = {'height': setHgh + 'px', 'font-size': '0'};
						jq(tn).css(cssProp);
						jq(bn).css(cssProp);
						break;
					case 'b':
						var tn = $t.$n(),
							tnHgh = tn.offsetHeight,
							tnOffsetTop = tn.offsetTop,
							setHgh = tnHgh - differHgh;
						jq(tn).css({'height': setHgh + 'px', 'font-size': '0'});
					}
				}
			}
		} : zk.$void,
		_restoreFontSize: function (n) {
			jq(n).css('font-size', jq(this.$n()).css('font-size'));
		},
		getText: function () {
			var $t = this.firstChild;
			switch (this.valign) {
			case 't':
				return $t.$n().innerHTML;
			case 'c':
				var $c = $t.nextSibling;
				return $c.$n().innerHTML;
			case 'b':
				var $b = this.lastChild; 
				return $b.$n().innerHTML;
			}
			return null;
		},
		getPureText: function () { //feature #26:Support copy/paste value to local Excel
			var $t = this.firstChild;
			switch (this.valign) {
			case 't':
				var tn = $t.$n();
				return tn.textContent || tn.innerText;
			case 'c':
				var $c = $t.nextSibling, cn = $c.$n();
				return cn.textContent || cn.innerText;
			case 'b':
				var $b = this.lastChild, bn = $b.$n();
				return bn.textContent || bn.innerText;
			}
			return null;
		},
		onSize: function () {
			var $t = this.firstChild,
				$c = $t.nextSibling,
				$b = this.lastChild;
			switch (this.valign) {
			case 't':
				this._initShowAndFlex([$t], false);
				break;
			case 'c':
				this._initShowAndFlex([$t, $b], false);
				break;
			case 'b':
				this._initShowAndFlex([$t], false);
				break;
			}
			this.fixFlexHeight();
		},
		/**
		 * Clean widget's text, set visible and flex it.
		 * @param widget list to show and flex
		 * @param boolean whether to clear text or not
		 */
		_initShowAndFlex: function (visWgts, clear) {
			var $t = this.firstChild,
				wgts = [$t, $t.nextSibling, this.lastChild],
				flex = false; 
			for (var i = 0; i < wgts.length; i++) {
				var w = wgts[i];
				if (visWgts.$contains(w)) {
					flex = true;
					if (clear)
						w.$n().innerHTML = '';
					w.$n().style.height = '';
					w.setVisible(true);
					w.setVflex(true);
				}	
			}
			if (flex)
				zWatch.fireDown('onSize', this);
		}
	});


	/**
	 * Cell inner container base on CSS flexbox for vertical alignment.
	 */
	zss.FlexboxInner = zk.$extends(zss.BaseInner, {
		_vtal: {'t': 'top', 'c': 'mid', 'b': 'btm'},
		_htal: {'l': 'left', 'c': 'center', 'r': 'right'},
		$init: function () {
			this.$supers('$init', arguments);
			this.appendChild(new Div());
			this.setWidth('100%');
			this.setHeight('100%');
		},
		setText: function (t, valign, halign) {
			var n = this.firstChild.$n(),
				cls = 'zscell-inner-';
			n.innerHTML = t;
			cls += this._vtal[valign] + '-' + this._htal[halign];
			jq(n.parentNode).attr('class', '').addClass(cls);
			this.onSize();
		},
		getText: function () {
			return this.firstChild.$n().innerHTML;
		},
		getPureText: function () { //feature #26: Support copy/paste value to local Excel
			var tn = this.firstChild.$n(); 
			return tn.textContent || tn.innerText;
		},
		onSize: zk.gecko ? function () {
			var n = this.$n();
			if (!n) {
				this.clearCache();
				n = this.$n();
			}
			jq(this.firstChild.$n()).css('width', jq(n).css('width'));
		} : zk.$void
	});
	
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
		ERROR_CELL = 5;

zss.Cell = zk.$extends(zk.Object, {
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
	$init: function (sheet, block, cmp, edit) {
		this.$supers('$init', arguments);
		this.id = cmp.id;
		this.sheetid = sheet.sheetid;
		this.sheet = sheet;
		this.block = block;
		this.comp = cmp;
		this.edit = edit;

		var $n = jq(cmp);
		this.r = zk.parseInt($n.attr('z.r'));
		this.c = zk.parseInt($n.attr('z.c'));
		this.defaultRowHgh = ('1' == $n.attr('z.drh'));
			
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
		if (valign)
			this.valign = valign;

		var inner = null,
			txtCmp = this.txtcomp = $n.children('DIV:first')[0],
			txt = this.txtcomp.innerHTML,
			vflex = this._vflex = sheet._wgt._cssFlex;/*support CSS flexbox or not; FF flexbox cause bug 340*/
		if (this.valign != "t" && txt) {
			inner = this._appendInner();
		}
		//is this cell has right border,
		//i must do some special processing when textoverfow when with rborder.
		this.rborder = $n.attr('z.rbo') == "t";

		if (!txt) {
			txt = "";
		} else {
			//TODO should i check space cell for better performance?
			/*var t = txt.trim();
			(txt.charCodeAt(0)==160 ){//160 is &nbsp;
			if(t.length ==0){
				
			}*/
		}

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
			if (inner) {
				//set text directly when DOM is ready
				if (inner.$n() && this.defaultRowHgh)
					inner.setText(txt, this.valign, this.halign, true);
				else //feel faster using insertSSInitLater
					sheet.insertSSInitLater(zss.Cell._processInnerText, this, inner, txt);
			}
			if (this.overflow = evalOverflow(this))
				sheet.addSSInitLater(zss.Cell._processOverflow, this);
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
	 * Append inner child for vertical alignment
	 * @param boolean support flexbox or not
	 */
	_appendInner: function () {
		var inner = null,
			txtCmp = this.txtcomp,
			ctmp = jq('<div id="ctmp"></div>');
		$(txtCmp).empty().append(ctmp);
		inner = this._inner = this._vflex ? new zss.FlexboxInner() : new zss.FlexInner();
		if (inner.onSize)
			this.onSize = function () {
				inner.onSize();
			};
		inner.replaceHTML(ctmp);
		inner.setSheet(this.sheet);
		return inner;
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
				zss.Cell._processOverflow(this);
				this.redoOverflow = false;
			}
		}
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

		var vtal = this.valign,
			hal =  this.halign,
			wrap = this.wrap;
		if (this._inner)
			this._inner.setText(txt, vtal, hal);
		else if (vtal == 't')
			this.txtcomp.innerHTML = txt;
		else if (txt) {
			this._appendInner();
			this._inner.setText(txt, vtal, hal);
		}
	},
	/**
	 * Invoke after append this cell to DOM tree
	 */
	afterAppend: zk.ie6_ ? function () {
		if (this._inner)
			this._inner.fixFlexHeight();
	} : zk.$void,
	cleanup: function () {
		this.invalid = true;
		this._updateHasTxt(false);
		if (this.overflowed && !this.block.invalid) {
			//remove the overlap relation , because the cell that is overlaped maybe not be destroied.
			zss.Cell._clearOverlapRelation(this);
		}
		var inner = this._inner;
		if (inner) {
			inner.unbind();
			this._inner = null;
		}
		if (this.comp) this.comp.ctrl = null;
		this.comp = this.txtcomp = this.sheet = this.overlapBy = 
			this.block = this._inner = this._vflex = this.lock = this.defaultRowHgh = null;
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
		
		var txtcmp = document.createElement("div");
		if (ist)
			txtcmp.style.cssText = ist;
		txtcmp.innerHTML = txt ? txt : "";
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
		cmp.style.cssText = !st ? '' : st;
		ctrl.txtcomp.style.cssText = !ist ? '' : ist;
		ctrl.halign = !hal ? "l" : hal;//default h align is left
		ctrl.valign = !vtal ? 't' : vtal;//default v align is top
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
	 * Invoke after CSS ready, set text and alignment
	 */
	_processInnerText: function (ctrl, cellInnerWgt, txt) {
		if (ctrl.invalid)
			return;
		
		cellInnerWgt.setText(txt, ctrl.valign, ctrl.halign, ctrl.defaultRowHgh);
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