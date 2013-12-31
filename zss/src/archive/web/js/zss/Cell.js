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
var WRAP_TEXT_CLASS = 'zscelltxt-wrap';
var NUM_CELL = 0,
	STR_CELL = 1,
	FORMULA_CELL = 2,
	BLANK_CELL = 3,
	BOOLEAN_CELL = 4,
	ERROR_CELL = 5,
	Cell = 
zss.Cell = zk.$extends(zk.Widget, {
	widgetName: 'Cell',
	/**
	 * Row index number
	 */
	//r
	/**
	 * Column index number
	 */
	//c
	/**
	 * Cell reference address
	 */
	ref: null,
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
	 * Indicate whether should invoke process overflow on cell or not.
	 * Process overflow when cell type is string, halign is left, no wrap, no merge
	 * 
	 * Currently, supports left aligned cell only
	 */
	overflow: false,
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
	 * Default: align bottom
	 */
	valign: 'b',
	/**
	 * Whether the text should be wrapped
	 * 
	 * Default: false
	 */
	wrap: false,
	/**
	 * The font size in point
	 */
	fontSize: 11,
	/**
	 * Whether listen onRowHeightChanged event or not 
	 * Currently, use only on IE6/IE7 for vertical align
	 * 
	 * Default: false
	 */
	//_listenRowHeightChanged: false,
	/**
	 * Whether listen sheet's onProcessOverflow event or not
	 * 
	 * Default: false
	 */
	//_listenProcessOverflow: false,
	$init: function (sheet, block, row, col, src) {
		this.$supers(zss.Cell, '$init', []);
		
		this.sheet = sheet;
		this.block = block;
		this.r = row;
		this.c = col;
		this.src = src;
		
		var	cellData = src.getRow(row).getCell(col),
			colHeader = src.columnHeaders[col],
			rowHeader = src.rowHeaders[row];
		this.text = cellData.text || '';
		if (colHeader && rowHeader) {
			this.ref = colHeader.t + rowHeader.t;
		}
		this.edit = cellData.editText ? cellData.editText : '';
		this.hastxt = !!this.text;
		this.zsw = src.getColumnWidthId(col);
		this.zsh = src.getRowHeightId(row);
		this.lock = cellData.lock;
		this.cellType = cellData.cellType;
		
		this.halign = cellData.halign;
		this.valign = cellData.valign;
		this.rborder = cellData.rightBorder;
		if (cellData.fontSize){
			this.fontSize = cellData.fontSize;
		}

		var mId = cellData.mergeId;
		if (mId) {
			var r = cellData.merge;
			this.merid = mId;
			this.merl = r.left;
			this.merr = r.right;
			this.mert = r.top;
			this.merb = r.bottom;
			this.mergeCls = cellData.mergeCls;
		}
		this.wrap = cellData.wrap;
		this.overflow = cellData.overflow;
		// ZSS-224: a overflow options for indicating more status in bitwise format
		// refer to Spreadsheet.java -> getCellAttr()
		this.overflowOpt = cellData.overflowOpt;  
		
		this.style = cellData.style;
		this.innerStyle = cellData.innerStyle;
	},
	getVerticalAlign: function () {
		switch (this.valign) {
		case 'b':
			return 'verticalAlignBottom';
		case 'c':
			return 'verticalAlignMiddle';
		case 't':
			return 'verticalAlignTop';
		}
	},
	getHorizontalAlign: function () {
		switch (this.halign) {
		case 'l':
			return 'horizontalAlignLeft';
		case 'c':
			return 'horizontalAlignCenter';
		case 'r':
			return 'horizontalAlignRight';
		}
	},
	getFontName: function () {
		var fn = jq(this.getTextNode()).css('font-family');
		if(fn){
			fn = fn.replace(/'/g,"");//replace "'" in some font that has Space
		}
		return fn;
	},
	/**
	 * Return cell font size in point
	 * 
	 * @return int font size
	 */
	getFontSize: function () {
		return this.fontSize;
	},
	isFontBold: function () {
		var b = jq(this.getTextNode()).css('font-weight');
		return b && (b == '700' || b == 'bold');
	},
	isFontItalic: function () {
		return jq(this.getTextNode()).css('font-style') == 'italic';
	},
	isFontUnderline: function () {
		var s = jq(this.$n('cave')).css('text-decoration');
		return s && s.indexOf('underline') >= 0;
	},
	isFontStrikeout: function () {
		var s = jq(this.$n('cave')).css('text-decoration');
		return s && s.indexOf('line-through') >= 0;
	},
	doClick_: function (evt) {
		//do nothing. eat the event.
	},
	doRightClick_: function (evt) {
		this.sheet._doMouserightclick(evt);
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
			data = this.src.getRow(r).getCell(c),
			format = data.formatText,
			st = this.style = data.style,
			ist = this.innerStyle = data.innerStyle,
			n = this.comp,
			overflow = data.overflow,
			cellType = data.cellType,
			txt = data.text,
			txtChd = txt != this.text,
			wrapChanged = this.wrap != data.wrap,
			processWrap = data.wrap || wrapChanged || (this.wrap && this.getText() != data.text),
			cave = this.$n('cave'),
			prevWidth = cave.style.width,
			fontSize = data.fontSize;
		if (fontSize) {
			this.fontSize = fontSize;
		}
		this.$n().style.cssText = st;
		cave.style.cssText = ist;
		if (prevWidth && (zk.ie6_ || zk.ie7_)) {//IE6/IE7 set overflow width at cave
			cave.style.width = prevWidth;
		}
		
		this.lock = data.lock;
		this.wrap = data.wrap;
		this.halign = data.halign;
		this.valign = data.valign;
		this.rborder = data.rightBorder;
		this.edit = data.editText;
		
		this._updateListenOverflow(overflow);
		this.setText(txt, false, wrapChanged); //when wrap changed, shall re-process overflow
		
		if (wrapChanged) {
			if (this.wrap) {
				jq(this.getTextNode()).addClass(WRAP_TEXT_CLASS);
			} else {
				jq(this.getTextNode()).removeClass(WRAP_TEXT_CLASS);
			}
		}
		
		if (this.overflow != overflow // overflow changed
			|| (this.overflow && txtChd)) { // already overflow and text changed
			var processedOverflow = false;
			if (this.overflow && !overflow) {
				this._clearOverflow();
				processedOverflow = true;
			}
			this.overflow = overflow;
			if (!processedOverflow)
				this._processOverflow();	
		}
		
		//trigger process overflow event when empty cell <-> string cell
		//overflow cells before this cell need to re-evaluate overflow
		if (this.cellType != cellType
			&& (this.cellType == STR_CELL || this.cellType == BLANK_CELL)) {//when cell type become empty, cells that before this cell have to re-process overflow
			this.sheet.triggerOverflowColumn_(this.r, this.c);
		}
		this.cellType = cellType;
		
		if (txtChd && processWrap)
			this._txtHgh = jq(this.getTextNode()).height();//cache txt height
		//merged cell won't change row height automatically
		if (this.cellType == STR_CELL && !this.merid && processWrap) {//must process wrap after set text
			this.parent.processWrapCell(this, true);
		}
	},
	/**
	 * Set overflow attribute and register listener or unregister onProcessOverflow listener base on overflow attribute
	 * @param boolean 
	 * @return boolean whether reset overflow attribute or not
	 */
	_updateListenOverflow: function (b) {
		var curr = !!this._listenProcessOverflow;
		if (curr != b) {
			this.sheet[curr ? 'unlisten' : 'listen']({onProcessOverflow: this.proxy(this._onProcessOverflow)});
			this._listenProcessOverflow = b;
			return true;
		}
		return false;
	},
	_updateHasTxt: function (bool) {
		this.hastxt = bool;
		
		var zIdx = this.getZIndex();
		jq(this.comp).css('z-index', bool ? "" : zIdx);
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
	},
	/**
	 * Returns the text node height. Shall invoke this method after CSS ready
	 * 
	 * @return int height
	 */
	getTextHeight: function () {
		var h = this._txtHgh;
		return h != undefined ? h : this._txtHgh = jq(this.getTextNode()).height();
	},
	_updateVerticalAlign: zk.ie6_ || zk.ie7_ ? function () {
		var	v = this.valign,
			text = this.text,
			cv = this.$n('cave');
		if (cv.style.display == 'none' || !text)
			return;
		var	$n = jq(cv),
			$t = jq(this.getTextNode());
		
		switch (v) {
		case 't':
			$t.css({'top': "0px", 'bottom': ''});
			break;
		case 'c':
			var ch = $n.height(),
				ich = $t.height();
			if (!ch || !ich)
				return;
			var	ah = (ch - ich) / 2,
				p = Math.ceil(ah * 100 / ch);
			if (p)
				$t.css({'top': p + "%", 'bottom': ''});
			break;
		case 'b':
			$t.css({'bottom': "0px", 'top': ''});	
			break;
		}
	} : zk.$void(),
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
		return this.getTextNode().innerHTML;
	},
	getPureText: function () { //feature #26: Support copy/paste value to local Excel\
		var n = this.getTextNode();
		return n.textContent || n.innerText;
	},
	_setText: function (txt) {
		if (!txt)
			txt = "";
		this.text = this.getTextNode().innerHTML = txt;
		
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
			style = this.domStyle_(),
			innerStyle = this.innerStyle;
		
		//IE6/IE7: vertical align need position:absolute;
		return '<div id="' + uid + '" class="' + this.getZclass() + '" zs.t="SCell" '
			+ (style ? 'style="' +  style + '"' : '') + '><div id="' + uid + '-cave" class="' +
			this._getInnerClass() + '" ' + (innerStyle ? 'style="' + innerStyle + '"' : '') + 
			'>' + '<div id="' + uid + '-real" class="zscelltxt-real '+(this.wrap?WRAP_TEXT_CLASS:'')+'">' + text + '</div>' + '</div></div>';
	},
	getZIndex: function () {
		if (zk.ie6_ || zk.ie7_)
			return this.cellType == BLANK_CELL ? -1 : 1;
		return this.text ? null : 1;
	},
	domStyle_: function () {
		var st = this.style;
		if (st) {
			return st;
		} else {
			var zIdx = this.getZIndex();
			if (zIdx){
				return 'z-index:' + zIdx + ';';
			}
		}
	},
	getTextNode: function () {
		return this.$n('real');
	},
	_clearOverflow: function () {
		jq(this.getTextNode()).css('width', '');//clear overflow
		jq(this.$n()).removeClass("zscell-overflow").removeClass("zscell-overflow-b");
	},
	_processOverflow: function () {
		// not implement in OSE
	},
	bind_: function (desktop, skipper, after) {
		this.$supers(zss.Cell, 'bind_', arguments);
		
		var n = this.comp = this.$n(),
			sheet = this.sheet;
		n.ctrl = this;
		this.cave = n.firstChild;


		if (this.cellType == BLANK_CELL) {//no need to process overflow and wrap
			return;
		}
		this._updateListenOverflow(this.overflow);
		
		if (!!this.text && (zk.ie6_ || zk.ie7_) && this.valign != 't') { // IE6 / IE7 doesn't support CSS vertical align
			this._updateListenRowHeightChanged(true);
			this._updateVerticalAlign();
		}

		// ZSS-224: skip process overflow according to the hint from server
		// it indicates that this cell's silbing isn't blank
		var skipOverflowOnBinding = (this.overflowOpt & 2) != 0; // skip overflow when initializing
		if (this.overflow && !skipOverflowOnBinding) {
			this._processOverflow(); // heavy duty
		}
		
		//merged cell won't change row height automatically
		if (this.cellType == STR_CELL && !this.merid && this.wrap) {
			//true indicate delay calcuate wrap height after CSS ready	
			this.parent.processWrapCell(this, true);
		}
	},
	unbind_: function () {
		this._updateListenOverflow(false);
		this._updateListenRowHeightChanged(false);
		
		this.comp = this.comp.ctrl = this.cave = this.sheet = this.overlapBy = this._listenRowHeightChanged =
		this.block = this.lock = null;
		
		this.$supers(zss.Cell, 'unbind_', arguments);
	},
	doMouseDown_: function (evt) {
		this.sheet._doMousedown(evt);
	},
	doMouseUp_: function (evt) {
		this.sheet._doMouseup(evt);
	},
	/**
	 * When cells after this cell changed, may effect this cell's overflow
	 */
	_onProcessOverflow: function (evt) {
		if (this.overflow) {
			var row = this.r,
				data = evt.data;
			if (data) {
				var rCol = data.col,
					tRow = data.tRow,
					bRow = data.bRow;
				if (this.c < data.col) {
					if ((tRow == undefined && bRow == undefined) || 
						(tRow && bRow && row >= tRow && row <= bRow)) {
						this._processOverflow(true);
					}
				}
			}
		}
	},
	//super//
	getZclass: function () {
		var cls = 'zscell',
			hId = this.zsh,
			wId = this.zsw,
			mCls = this.mergeCls;
		if (hId)
			cls += (' zshi' + hId);
		if (wId)
			cls += (' zsw' + wId);
		if (mCls)
			cls += (' ' + mCls);
		return cls;
	},
	_getInnerClass: function () {
		var cls = 'zscelltxt',
			hId = this.zsh,
			wId = this.zsw;
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
		if (zsw) {
			this.zsw = zsw;
			jq(this.comp).addClass("zsw" + zsw);
			jq(this.cave).addClass("zswi" + zsw);
		}
	},
	/**
	 * Sets the height position index
	 * @param int zsh the height position index
	 */
	appendZSH: function (zsh) {
		if (zsh) {
			this.zsh = zsh;
			jq(this.comp).addClass("zshi" + zsh);
			jq(this.cave).addClass("zshi" + zsh);	
		}
	},
	/**
	 * Sets the column index of the cell
	 * @param int column index
	 */
	resetColumnIndex: function (newcol) {
		var	src = this.src;
		this.ref = src.columnHeaders[newcol].t + src.rowHeaders[this.r].t;
		this.c = newcol;
	},
	/**
	 * Sets the row index of the cell
	 * @param int row index
	 */
	resetRowIndex: function (newrow) {
		var	src = this.src;
		this.ref = src.columnHeaders[this.c].t + src.rowHeaders[newrow].t;
		this.r = newrow;
	}
});
})();