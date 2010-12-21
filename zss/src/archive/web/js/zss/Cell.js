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

zss.Cell = zk.$extends(zk.Object, {
	$init: function (sheet, block, cmp, edit) {
		this.$supers('$init', arguments);
		this.id = cmp.id;
		this.sheetid = sheet.sheetid;
		this.sheet = sheet;
		this.block = block;
		this.comp = cmp;
		this.overby = null;
		this.overhead = false;//overhead, this cell is a overflow's left cell
		this.edit = edit;
		
		var local = this;
		this.r = zk.parseInt(jq(cmp).attr('z.r'));
		this.c = zk.parseInt(jq(cmp).attr('z.c'));
			
		this.zsw = jq(cmp).attr('z.zsw');

		if (this.zsw)
			this.zsw = zk.parseInt(this.zsw);

		this.zsh = jq(cmp).attr('z.zsh');

		if (this.zsh)
			this.zsh = zk.parseInt(this.zsh);
		
		this.wrap = (jq(cmp).attr('z.wrap') == "t") ? true : false;
		this.halign = jq(cmp).attr('z.hal');
		if (!this.halign)
			this.halign = "l";//default h align is left;
		//is this cell has right border,
		//i must do some special processing when textoverfow when with rborder.
		this.rborder = (jq(cmp).attr('z.rbo') == "t") ? true : false;

		this.txtcomp = jq(cmp).children('DIV:first')[0];
			
		var txt = this.txtcomp.textContent || this.txtcomp.innerText;
		if (!txt) {
			txt = "";
		} else {
			//TODO should i check space cell for better performance?
			/*var t = txt.trim();
			(txt.charCodeAt(0)==160 ){//160 is &nbsp;
			if(t.length ==0){
				
			}*/
		}

		if (txt == '')
			this._updateHasTxt(false);
		else {
			this._updateHasTxt(true);
			if (sheet.config.textOverflow && !this.wrap)
				sheet.addSSInitLater(zss.Cell._processOverflow, this);
		}
		
		//process merge.
		var mr = jq(cmp).attr('z.merr'),
			mid = jq(cmp).attr('z.merid'),
			ml = jq(cmp).attr('z.merl');
		
		if (zkS.t(mid)) {
			this.merr = zk.parseInt(mr); 
			this.merid = zk.parseInt(mid);
			this.merl = zk.parseInt(ml); 
		}
		cmp.ctrl = this;
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

		this._updateHasTxt(txt != "" ? true : false);

		this.txtcomp.innerHTML = txt;
		if (this.sheet.config.textOverflow)
			zss.Cell._processOverflow(this);
	},
	_setText: function (txt) {
		if (!txt)
			txt = "";
		this.txtcomp.innerHTML = txt;
	},
	cleanup: function () {
		this.invalid = true;
		this._updateHasTxt(false);
		if (this.overhead && !this.block.invalid) {
			//remove the overby relation , because the cell that is overbyed maybe not be destroied.
			zss.Cell._clearOverbyRelation(this);
			this.block._overflowcell.$remove(this);	
		}
		if (this.comp) this.comp.ctrl = null;
		this.comp = this.txtcomp = this.sheet = this.overby = this.block = null;
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
		this.c = newcol;
		jq(this.comp).attr("z.c", newcol);
	},
	/**
	 * Sets the row index of the cell
	 * @param int row index
	 */
	resetRowIndex: function (newrow) {
		this.r = newrow;
		jq(this.comp).attr("z.r", newrow);
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
			edit = parm.edit,
			st = parm.st,//style
			ist = parm.ist,//style of inner div
			wrap = parm.wrap,
			hal = parm.hal,
			rbo = parm.rbo,
			merr = parm.merr,
			merid = parm.merid,
			merl = parm.merl,
			zsw = parm.zsw,
			zsh = parm.zsh,
			cmp = document.createElement("div");

		jq(cmp).attr("zs.t", "SCell").attr("z.c", col).attr("z.r", row);
		
		if (zsw)
			jq(cmp).attr("z.zsw", zsw);

		if (zsh)
			jq(cmp).attr("z.zsh", zsh);
		
		if (wrap) jq(cmp).attr("z.wrap", "t");
		jq(cmp).attr("z.hal", (hal ? hal : "l"));
		if (rbo) jq(cmp).attr("z.rbo", "t");
		
		if (zkS.t(merid)) {
			jq(cmp).attr("z.merr", merr).attr("z.merid", merid).attr("z.merl", merl).addClass(merl == col ? "zsmerge" + merid : "zsmergee");
		}
		
		if (st)
			cmp.style.cssText = st;
	
		if (!txt)
			txt = "";
		
		var txtcmp = document.createElement("div");
		if (ist)
			txtcmp.style.cssText = ist;

		txtcmp.innerHTML = txt;
		cmp.appendChild(txtcmp);
		
		var sclazz = "zscell" + (zsw ? " zsw" + zsw : "") + (zsh ? " zshi" + zsh : "");
		jq(cmp).addClass(sclazz);
		
		sclazz = "zscelltxt" + (zsw ? " zswi" + zsw : "") + (zsh ? " zshi" + zsh : "");
		jq(txtcmp).addClass(sclazz);

		return new zss.Cell(sheet, block, cmp, edit);
	},
	/**
	 * Update cell's text and style
	 * @param zss.Cell ctrl
	 * @param map parameter setting
	 */
	updateCell: function (ctrl, parm) {
		var txt = parm.txt,
			st = parm.st,//style
			ist = parm.ist,//style of inner div
			wrap = parm.wrap,
			hal = parm.hal,
			rbo = parm.rbo;

		var cmp = ctrl.comp;
		
		if (!wrap)
			wrap = false;
		jq(cmp).attr("z.wrap", wrap);	
		ctrl.wrap = wrap;
		var txtcmp = ctrl.txtcomp;
		
		if (!st)
			st = "";
		if (!ist)
			ist = "";
		cmp.style.cssText = st;
		txtcmp.style.cssText = ist;
		
		ctrl.halign = hal;
		if (!ctrl.halign)
			ctrl.halign = "l";//default h align is left;
		
		ctrl.rborder = (rbo == true);

		ctrl.setText(txt);
		ctrl.edit = parm.edit;
	},
	_clearOverbyRelation: function (ctrl) {
		var cmp = ctrl.comp,
			sheet = ctrl.sheet;
		if (ctrl.overhead && !sheet.invalid) {
			var next = jq(cmp).next('DIV')[0],
				nctrl;
			while (next) {
				nctrl = next.ctrl;
				if (nctrl && nctrl.overby == ctrl) {
					nctrl.overby = null;
					next = jq(next).next('DIV')[0];
					continue;	
				}
				break;
			}
		}
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
		//1995735,force IE recaluate some thing first
		if (zk.ie && txtcmp.clientWidth) {}

		var sw = txtcmp.scrollWidth,
			cw = cmp.clientWidth - 2 * sheet.cellPad,
			oldoverby = ctrl.overby,
			overflow = (ctrl.hastxt && ctrl.halign == "l" && !ctrl.wrap  && !zkS.t(ctrl.merid)  && sw > cw); 
		
		if (overflow) {//process condition, halign is left, no wrap
			if (!ctrl.overhead && !block._overflowcell.$contains(ctrl))
				block._overflowcell.push(ctrl);

			ctrl.overhead = true;
			txtcmp.style.position = zk.ie || zk.safari ? "absolute" : "relative";

			var w = cmp.clientWidth,
				prev = cmp,
				next = jq(cmp).next('DIV')[0],
				count = 0;
			while (next) {
				if (!next.ctrl) {//next no initial yet, then delay and process again.
					zkS.addDelayBatch(zss.Cell._processOverflow, false, ctrl);
					return;
				}
				next.ctrl.overby = ctrl;
				if (next.ctrl.hastxt) {
					jq(prev).removeClass(prev.ctrl.rborder ? "zscell-over-b" : "zscell-over");
					break;
				}
				if (prev.ctrl.rborder) {
					jq(prev).removeClass("zscell-over").addClass("zscell-over-b");
				} else {
					jq(prev).removeClass("zscell-over-b").addClass("zscell-over");
				}
				
				var ow = next.offsetWidth;
				w = w + ow;
				if (w > sw) {
					jq(next).removeClass(next.ctrl.rborder ? "zscell-over-b" : "zscell-over");
					
					//keep remove next overflow if it is over by this cell
					next = jq(next).next('DIV')[0];//next.nextSibling;
					while (next) {
						if (ctrl != next.ctrl.overby)
							break;
						jq(next).removeClass(next.ctrl.rborder ? "zscell-over-b" : "zscell-over");

						next.ctrl.overby = null;
						next = jq(next).next('DIV')[0];
					}
					break;
				}
				prev = next;
				next = jq(next).next('DIV')[0];
			}
			w =  w - sheet.cellPad;
			if (w < 0) w = 0;// bug, when some column or row is invisible, the w might be samlle then zero
			jq(txtcmp).css('width', jq.px0(w));
		} else {
			if (ctrl.overhead)
				block._overflowcell.$remove(ctrl);	
			ctrl.overhead = false;
			jq(cmp).removeClass(cmp.ctrl.rborder ? "zscell-over-b" : "zscell-over");

			var next = jq(cmp).next('DIV')[0];
			while (next) {
				//next no initial yet skip to check
				if(!next.ctrl || 
					ctrl != next.ctrl.overby && (oldoverby ? oldoverby != next.ctrl.overby : true))
					break;
				jq(next).removeClass(next.ctrl.rborder ? "zscell-over-b" : "zscell-over");

				next.ctrl.overby = null;
				next = jq(next).next('DIV')[0];
			}
		}
		
		if (oldoverby)
			zss.Cell._processOverflow(oldoverby);
	}
});