/* HeadCtrl.js

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
	function _ignoresizing (dg, pt, evt) {
		var ctrl = dg.control;
		ctrl.draging = true;
		zss.SSheetCtrl._curr(ctrl).dp.stopEditing("refocus");
		
		ctrl.drag.start = [pt[0], pt[1]];
		
		return ctrl.drag._fixstart = false;
	}
	
	function _startDrag (dg, evt) {
		dg.control.sheet.headerdrag = true;
	}

	function _endDrag (dg, evt) {
		var ctrl = dg.control,
			sheet = ctrl.sheet,
			cmp = ctrl.comp;
		ctrl.draging = sheet.headerdrag = false;
		
		if (ctrl.type == zss.Header.HOR) {
			var fw,
				offset = ctrl.drag.last[0] - ctrl.drag.start[0];
			fw = ctrl.orgsize + offset;
			if (fw < ctrl.minHWidth) fw = ctrl.minHWidth;
			
			jq(ctrl.comp).css('width', '');
			jq(ctrl.icomp).css('width', '');
			sheet._setColumnWidth(ctrl.index, fw, true, true);
		} else {
			var fh,
				offset = ctrl.drag.last[1] - ctrl.drag.start[1];
			fh = ctrl.orgsize + offset;
			if (fh < ctrl.minVHeight) fh = ctrl.minVHeight;

			jq(ctrl.comp).css({'height': '', 'line-height': ''});
			jq(ctrl.icomp).css({'height': '', 'line-height': ''});

			sheet._setRowHeight(ctrl.index, fh, true, true);
		}

		//gain focus and reallocate mark , then show it, 
		//don't move the focus because the cell maybe doens't exist in block anymore.		
		sheet.dp._gainFocus(true);
		
		var pos = sheet.getLastFocus(),
			ls = sheet.getLastSelection();

		sheet.moveCellFocus(pos.row, pos.column);
		sheet.moveCellSelection(ls.left, ls.top, ls.right, ls.bottom);
		ctrl.orgsize = null;
	}

	
	function _snap (dg, pt) {
		var ctrl = dg.control,
			drag = ctrl.drag,
			last = [pt[0], pt[1]],
			cmp = ctrl.comp;

		if (!drag._fixstart) {
			drag.start[0] -= drag.offset[0];
			drag.start[1] -= drag.offset[1];
			drag._fixstart = true;

			//when column size is 0, the head will set to display none, so, we remove the style here
			//see also zkSSheetCtrl#setColumnWidth
			//var name = "#"+this.sheetid;
			//zcss.setRule(name+" div.zscw"+this.index,"display","",true,this.sheetid+"-sheet");
			ctrl._processDrag(false);
		}

		if (ctrl.type == zss.Header.HOR) {
			if (ctrl.orgsize == null) 
				ctrl.orgsize = zk(cmp).offsetWidth();
			
			var off = ctrl.drag.start[0] - pt[0],
				maxoff = ctrl.orgsize - ctrl.minHWidth;
				
			if (maxoff < 0) maxoff = 0;
			
			last = off >= maxoff ? [ctrl.drag.start[0] - maxoff, 0] : [pt[0], 0];
			ctrl.drag.last = [last[0], last[1]];
				
			if (zk.opera) {//In opera , i must add head position to get correct left position
				var bcmp = cmp.ctrl.bcomp;
				last[0] += (bcmp.offsetLeft+bcmp.offsetWidth);
			}
			
			//set size of column right now, but it will fail in opera
			if (!zk.opera) {
				var w, wi,
					offset = last[0] - ctrl.drag.start[0];
				wi = w = ctrl.orgsize + offset;
				if(w < ctrl.minHWidth) w = ctrl.minHWidth;
				
				//for firfox -mox-boz-size, offsetWidth is the width
				if (zk.gecko)
					wi = zk(cmp).revisedWidth(w); 
				else 
					w = wi = zk(cmp).revisedWidth(w);
					
				jq(cmp).css('width', jq.px(w));
				jq(ctrl.icomp).css('width', jq.px(wi));
			}
		} else {
			if (ctrl.orgsize == null)
				ctrl.orgsize = zk(cmp).offsetHeight();

			var off = ctrl.drag.start[1] - pt[1],
				maxoff = ctrl.orgsize - ctrl.minVHeight;

			if (maxoff < 0) maxoff =0;
			last = off >= maxoff ? [0, ctrl.drag.start[1] - maxoff] : last = [0, pt[1]];

			ctrl.drag.last = [last[0], last[1]];
			//set size of row right now, but it will fail in opera
			if (!zk.opera) {
				var h,
					offset = last[1] - ctrl.drag.start[1];
				h = ctrl.orgsize + offset;
				if (h < ctrl.minVHeight) h= ctrl.minVHeight;
				if (h == 0)//prevent minus value.
					h = 1;

				jq(cmp).css({'height': jq.px(h - 1), 'line-height': jq.px(h - 1)});
				jq(ctrl.icomp).css({'height': jq.px(h - 1), 'line-height': jq.px(h - 1)});
			}
		}
		return last;
	}
	
	/**
	 * Create a dom element to be drag instead of drag the element itself
	 */
	function _ghosting (dg, ofs, evt) {
		var drag = dg.control,
			bcmp = drag.ibcomp,
			height = bcmp.offsetHeight/2 + 1,//3;
			width = bcmp.offsetWidth/2 + 1,//3;
			top = ofs[1],
			left = ofs[0],
			spcomp = zss.SSheetCtrl._curr(drag).sp.comp;

		if (drag.type == zss.Header.HOR) {
			var w = zk(spcomp).offsetWidth(),
				barHeight = (spcomp.scrollWidth - w <= 0) ? 0 : zss.Spreadsheet.scrollWidth,
			height = zk(spcomp).offsetHeight() - barHeight;
		} else {
			var h = zk(spcomp).offsetHeight(),
				barWidth = (spcomp.scrollHeight - h <= 0) ? 0 : zss.Spreadsheet.scrollWidth;
			width = zk(spcomp).offsetWidth() - barWidth;
		}

		var html = ['<div id="zk_sghost" style="background:#AAA;position:absolute;top:', top, 'px;left:',
		            left, 'px;width:', width, 'px;height:', height, 'px;"></div>'].join('');
		jq(document.body).append(html);
		
		return drag.element = jq('#zk_sghost')[0];
	}
/**
 * Header represent row/column header of the spreadsheet.
 */
zss.Header = zk.$extends(zk.Object, {
	draging: false,
	$init: function (sheet, header, boundary, index, type) {
		this.$supers('$init', arguments);
		this.id = header.id;
		this.sheetid = sheet.id;
		this.sheet = sheet;
		
		header.ctrl = this;
		this.comp = header;
		this.minHWidth = this.minVHeight = 0;
		this.type = type;
		this.index = index;
		this.orgsize = null;
		
		if (type == zss.Header.HOR) {
			this.zsw = jq(header).attr('z.zsw');
			if (this.zsw)
				this.zsw = zk.parseInt(this.zsw);
		} else {
			this.zsh = jq(header).attr('z.zsh');
			if (this.zsh) 
				this.zsh = zk.parseInt(this.zsh);
		}
		
		boundary.ctrlref = this;
		this.bcomp = boundary;
		this.ibcomp = jq(boundary).children("DIV:first")[0];
		this.icomp = jq(header).children("DIV:first")[0];
	},
	/**
	 * Returns the header node
	 * @return DOMElement
	 */
	getHeaderNode: function () {
		return this.comp;
	},
	cleanup: function () {
		if (this.comp) this.comp.ctrl = null;

		if (this.drag) {
			this.drag.destroy();
			this.drag = null;
		}
		this.cells = this.comp = this.bcomp.ctrlref = this.bcomp = 
			this.ibcomp = this.icomp = this.sheet = null;
	},
	/**
	 * Sets the width position index
	 * @param int the width position index
	 */
	appendZSW: function (zsw) {
		this.zsw = zsw;

		jq(this.comp).attr("z.zsw", zsw).addClass("zsw" + zsw);
		jq(this.icomp).addClass("zswi" + zsw);
	},
	/**
	 * Sets the height position index
	 * @param int the height position index
	 */
	appendZSH: function (zsh) {
		this.zsh = zsh;

		jq(this.comp).attr("z.zsh", zsh).addClass("zslh" + zsh);
		jq(this.icomp).addClass("zslh" + zsh);
	},
	/**
	 * Reset info
	 * @param int newindex new index
	 * @param string newnm title of the header
	 */
	resetInfo: function (newindex, newnm) {
		this.index = newindex;
		if (this.type == zss.Header.HOR) {
			jq(this.comp).attr("z.c", newindex);
			jq(this.icomp).text(newnm);
		} else {
			jq(this.comp).attr("z.r", newindex);
			jq(this.icomp).text(newnm);
		}
	},
	_processDrag: function (show) {
		if (this.sheet.isDragging()) return;//don't care if dragging
		var ibcmp = this.ibcomp;
		if (this.type == zss.Header.HOR)	
			jq(ibcmp)[show && !zss.Header.draging ? 'addClass' : 'removeClass']("zshbouni-over");
		else
			jq(ibcmp)[show && !zss.Header.draging ? 'addClass' : 'removeClass']("zsvbouni-over");
		
		if (!this.drag) {
			var local = this;
			this.drag = new zk.Draggable(this, ibcmp, {
				//revert: true,
				constraint: (local.type == zss.Header.HOR) ? "horizontal" : "vertical",
				ghosting: _ghosting,
				snap: _snap,
				ignoredrag: _ignoresizing,
				starteffect: _startDrag,
				endeffect: _endDrag
			});
		}
	}
}, {
	HOR: "H",
	VER: "V",
	/**
	 * Create header node
	 */
	createComp: function (sheet, parm) {
		var type = parm.type,
			index = parm.ix,
			name = parm.nm,
			zsw = parm.zsw,
			zsh = parm.zsh;
		if (type == zss.Header.HOR) {
			//create column
			var uuid = sheet._wgt.uuid,
				html = ['<div z.c="', index,'" zs.t="STheader" class="zstopcell', zsw ? ' zsw' + zsw : '', '" ',
					zkS.t(zsw) ? 'z.zsw="' + zsw + '"' : '' ,'><div class="', 'zstopcelltxt ', zsw ? ' zswi' + zsw : '', '"',
					'>', name, '</div></div><div  class="zshboun"><div class="zshbouni" zs.t="SBoun"></div></div>'].join('');
				$n = jq(html),
				cmp = $n[0],
				bcmp = $n[1];
			return new zss.Header(sheet, cmp, bcmp, index, type);
			
		} else if (type == zss.Header.VER) {
			//create row
			var html = ['<div z.r="', index, '" zs.t="SLheader" class="zsleftcell zsrow',
			            (zsh ? " zslh" + zsh : ''), '"', zkS.t(zsh) ? ' z.zsh="' + zsh + '"' : '', '><div class="zsleftcelltxt', 
			            (zsh ? " zslh" + zsh : ''), '">', name,'</div></div><div class="zsvboun"><div class="zsvbouni" zs.t="SBoun"></div></div>'].join(''),
			    $n = jq(html),
			    cmp = $n[0],
			    bcmp = $n[1];

			return new zss.Header(sheet, cmp, bcmp, index, type);
		}
	}
});
})();