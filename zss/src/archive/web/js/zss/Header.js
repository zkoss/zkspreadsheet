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
		var pt0 = pt[0] - (dg._unhide && ctrl.type == zss.Header.HOR ? 6 : 0),
			pt1 = pt[1] - (dg._unhide && ctrl.type != zss.Header.HOR ? 6 : 0);
		
		dg.start = [pt0, pt1];
		return dg._fixstart = false;
	}
	
	function _startDrag (dg, evt) {
		dg.control.sheet.headerdrag = true;
	}

	function _endDrag (dg, evt) {
		var ctrl = dg.control,
			sheet = ctrl.sheet,
			cp = sheet.cp,
			rng = sheet.getLastSelection(),
			cousin = ctrl.cousin,
			idx = ctrl.index,
			type = sheet.selType;
		
		ctrl.draging = sheet.headerdrag = false;
		
		if (ctrl.type == zss.Header.HOR) {
			var offset = dg.last[0] - dg.start[0],
				fw = ctrl.orgsize + offset,
				left = rng.left,
				right = rng.right;
			if (fw < ctrl.minHWidth) fw = ctrl.minHWidth;
			if (type == zss.SelDrag.SELCOL && left <= idx && idx <= right) {
				sheet._setColumnsWidth(left, right, fw, true, true, dg._unhide? false : undefined); //undefined means depends on fw
			} else
				sheet._setColumnWidth(idx, fw, true, true, dg._unhide? false : undefined); //undefined means depends on fw
		} else {
			var offset = dg.last[1] - dg.start[1],
				fh = ctrl.orgsize + offset,
				top = rng.top,
				bottom = rng.bottom;
			if (fh < ctrl.minVHeight) fh = ctrl.minVHeight;
			if (type == zss.SelDrag.SELROW && top <= idx && idx <= bottom) {
				sheet._setRowsHeight(top, bottom, fh, true, true, dg._unhide? false : undefined); //undefined means depends on fh
			} else
				sheet._setRowHeight(idx, fh, true, true, dg._unhide? false : undefined); //undefined means depends on fh
		}
		//clear header style
		jq(ctrl.comp).removeAttr('style');
		jq(ctrl.icomp).removeAttr('style');
		if (cousin) {
			jq(cousin.comp).removeAttr('style');
			jq(cousin.icomp).removeAttr('style');
			jq(cp.comp).removeAttr('style');
		}

		//gain focus and reallocate mark , then show it, 
		//don't move the focus because the cell maybe doens't exist in block anymore.		
		sheet.dp._gainFocus(true);
		
		var pos = sheet.getLastFocus(),
			ls = sheet.getLastSelection(); //must re-fetch again

		sheet.moveCellFocus(pos.row, pos.column);
		sheet.moveCellSelection(ls.left, ls.top, ls.right, ls.bottom);
		ctrl.orgsize = null;
	}

	
	function _snap (dg, pt) {
		var ctrl = dg.control,
			last = [pt[0], pt[1]],
			cmp = ctrl.comp,
			icmp = ctrl.icomp,
			sheet = ctrl.sheet,
			cp = sheet.cp,
			cousin = ctrl.cousin;
		if (!dg._fixstart) {
			dg.start[0] -= dg.offset[0];
			dg.start[1] -= dg.offset[1];
			dg._fixstart = true;

			//when column size is 0, the head will set to display none, so, we remove the style here
			//see also zkSSheetCtrl#setColumnWidth
			//var name = "#"+this.sheetid;
			//zcss.setRule(name+" div.zscw"+this.index,"display","",true,this.sheetid+"-sheet");
			ctrl._processDrag(false, dg._unhide);
		}

		if (ctrl.type == zss.Header.HOR) {
			if (ctrl.orgsize == null) 
				ctrl.orgsize = zk(cmp).offsetWidth();
			
			var off = dg.start[0] - pt[0],
				maxoff = ctrl.orgsize - ctrl.minHWidth;
			if (maxoff < 0) maxoff = 0;
			
			last = off >= maxoff ? [dg.start[0] - maxoff, 0] : [pt[0], 0];
			dg.last = [last[0], last[1]];
				
			if (zk.opera) {//In opera , i must add head position to get correct left position
				var bcmp = ctrl.bcomp;
				last[0] += (bcmp.offsetLeft+bcmp.offsetWidth);
			}
			
			//set size of column right now, but it will fail in opera
			if (!zk.opera) {
				var w, wi, ow,
					offset = last[0] - dg.start[0];
				ow = wi = w = ctrl.orgsize + offset;
				if(w < ctrl.minHWidth) w = ctrl.minHWidth;
				
				//for firfox -moz-box-size, offsetWidth is the width
				jq(cmp).removeAttr('style');
				if (ow < 5) {
					jq(cmp).css('padding', 0);
					if (ow <= 0)
						jq(cmp).css('border-right', 'none');
				}
				if (zk.gecko)
					wi = zk(cmp).revisedWidth(w); 
				else 
					w = wi = zk(cmp).revisedWidth(w);
				
				jq(cmp).css('width', jq.px0(w));
				jq(icmp).css('width', jq.px0(wi));
				if (cousin) {
					jq(cousin.comp).removeAttr('style');
					if (ow < 5) {
						jq(cousin.comp).css('padding', 0);
						if (ow <= 0)
							jq(cousin.comp).css('border-right', 'none');
					}
					jq(cousin.comp).css('width', jq.px0(w));
					jq(cousin.icomp).css('width', jq.px0(wi));
					jq(cp.comp).css('width', jq.px0(cp._cornerWidth() + w - ctrl.orgsize + (!zk.gecko ? ow >= 5 ? 5 : ow > 0 ? 1 : 0 : 0)));
				}
			}
		} else {
			if (ctrl.orgsize == null)
				ctrl.orgsize = zk(cmp).offsetHeight();

			var off = dg.start[1] - pt[1],
				maxoff = ctrl.orgsize - ctrl.minVHeight;

			if (maxoff < 0) maxoff =0;
			last = off >= maxoff ? [0, dg.start[1] - maxoff] : last = [0, pt[1]];

			dg.last = [last[0], last[1]];
			//set size of row right now, but it will fail in opera
			if (!zk.opera) {
				var h, oh,
					offset = last[1] - dg.start[1];
				oh = h = ctrl.orgsize + offset;
				if (h < ctrl.minVHeight) h= ctrl.minVHeight;

				jq(cmp).removeAttr('style');
				if (oh < 5) {
					jq(cmp).css('padding', 0);
					if (oh <= 0)
						jq(cmp).css('border-bottom', 'none');
				}
				jq(cmp).css({'height': jq.px0(h - 1), 'line-height': jq.px0(h - 1)});
				jq(icmp).css({'height': jq.px0(h - 1), 'line-height': jq.px0(h - 1)});
				if (cousin) {
					jq(cousin.comp).removeAttr('style');
					if (oh < 5) {
						jq(cousin.comp).css('padding', 0);
						if (oh <= 0)
							jq(cousin.comp).css('border-bottom', 'none');
					}
					jq(cousin.comp).css({'height': jq.px0(h - 1), 'line-height': jq.px0(h - 1)});
					jq(cousin.icomp).css({'height': jq.px0(h - 1), 'line-height': jq.px0(h - 1)});
					jq(cp.comp).css('height', jq.px0(cp._cornerHeight() + h - ctrl.orgsize));
				}
			}
		}
		return last;
	}
	
	/**
	 * Create a dom element to be drag instead of drag the element itself
	 */
	function _ghosting (dg, ofs, evt) {
		var htmlborder,
			ctrl = dg.control,
			bcmp = ctrl.ibcomp,
			bcmpw = bcmp.offsetWidth,
			bcmph = bcmp.offsetHeight,
			height = zk.ie ? 2 : 3,
			width = zk.ie ? 2 : 3,
			top = ofs[1] + bcmph,
			left = ofs[0] + bcmpw,
			spcomp = zss.SSheetCtrl._curr(ctrl).sp.comp;

		if (ctrl.type == zss.Header.HOR) {
			var w = zk(spcomp).offsetWidth(),
				barHeight = (spcomp.scrollWidth - w <= 0) ? 0 : zss.Spreadsheet.scrollWidth,
			height = zk(spcomp).offsetHeight() - barHeight - bcmph;
			htmlborder = 'border-right:thin dotted #000;';
		} else {
			var h = zk(spcomp).offsetHeight(),
				barWidth = (spcomp.scrollHeight - h <= 0) ? 0 : zss.Spreadsheet.scrollWidth;
			width = zk(spcomp).offsetWidth() - barWidth - bcmpw;
			htmlborder = 'border-bottom:thin dotted #000;';
		}
		
		if (jq('#zk_sghost')) //if exists, remove it first
			jq('#zk_sghost').remove();

		var html = ['<div id="zk_sghost" style="font-size:0;line-height:0px;', htmlborder, 'position:absolute;top:', top, 'px;left:',
	            left, 'px;width:', width, 'px;height:', height, 'px;z-index:20"></div>'].join('');
		jq(document.body).append(html);
		
		return ctrl.element = jq('#zk_sghost')[0];
	}
/**
 * Header represent row/column header of the spreadsheet.
 */
zss.Header = zk.$extends(zk.Object, {
	draging: false,
	$init: function (sheet, header, boundary, index, type, cousin) {
		this.$supers('$init', arguments);
		this.id = header.id;
		this.sheetid = sheet.id;
		this.sheet = sheet;
		//cornerPanel Header link to sheet cousin Header if any
		if (cousin) {
			this.cousin = cousin;
			cousin.cousin = this;
		}
		
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
		this.ibcomp2 = jq(boundary).children("DIV:last")[0];
		if (this.ibcomp2 == this.ibcomp) 
			delete this.ibcomp2;
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
		this.cousin = this.cells = this.comp = this.bcomp.ctrlref = this.bcomp = 
			this.ibcomp = this.icomp = this.sheet = null;
		if (this.ibcomp2)
			delete this.ibcomp2;
	},
	/**
	 * Setup column header per the new width or whether unhide the column header
	 */
	setColumnHeader: function (hidden) {
		var sheet = this.sheet,
			cmp = this.comp,
			icmp = this.icomp;
		
		jq(cmp).css('width', '');
		jq(icmp).css('width', '');

		//hidden so need the extra boundary to "unhide" the column 
		if (hidden) {
			if (!this.ibcomp2) //insert the SBoun2
				this.ibcomp2 = jq(this.ibcomp).after('<div class="zshbounw" zs.t="SBoun2"></div>').next()[0];
			//hide the sizing boundary to avoid affect sizing bounary of left side header
			jq(this.ibcomp).css('visibility', 'hidden');
		} else if (this.ibcomp2) {//if not hidden, must remove extra "unhide" boudnary if exist
			jq(this.ibcomp2).remove();
			jq(this.ibcomp).css('visibility', '');
			delete this.ibcomp2;
		}
	},
	
	/**
	 * Setup row header per the new height or whether unhide the row header
	 */
	setRowHeader: function (hidden) {
		var sheet = this.sheet,
		cmp = this.comp,
		icmp = this.icomp;
		
		jq(cmp).css({'height': '', 'line-height': ''});
		jq(icmp).css({'height': '', 'line-height': ''});

		//hidden so need the extra boundary to "unhide" the row
		if (hidden) {
			if (!this.ibcomp2) //insert the SBoun2
				this.ibcomp2 = jq(this.ibcomp).after('<div class="zsvbounw" zs.t="SBoun2"></div>').next()[0];
			//hide the sizing boundary so will not disturb bottom headers
			jq(this.ibcomp).css('visibility', 'hidden');
		} else if (this.ibcomp2) {//if not hidden, must remove extra "unhide" boudnary if exist
			jq(this.ibcomp2).remove();
			jq(this.ibcomp).css('visibility', '');
			delete this.ibcomp2;
		}
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
	_processDrag: function (show, unhide) { //sizing column
		if (this.sheet.isDragging())
			return;//don't care if dragging
		
		if (show && this.drag && this.drag._unhide != unhide) {
			if (this.draging) { //drag direct from unhide thumb to size thumb, shall ignore such case 
				return;
			}
			delete this.drag;
		}
		
		var ibcmp = unhide ? this.ibcomp2 : this.ibcomp,
			ibcmpcls = this.type == zss.Header.HOR ? 
				(unhide ? "zshbounw-over" : "zshbouni-over") : (unhide ? "zsvbounw-over" : "zsvbouni-over");
		jq(ibcmp)[show && !zss.Header.draging ? 'addClass' : 'removeClass'](ibcmpcls);
		 
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
			this.drag._unhide = unhide; //indicate that this dragging is for unhide
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
			zsh = parm.zsh,
			hidden = parm.hn;
		if (type == zss.Header.HOR) {
			//create column
			var uuid = sheet._wgt.uuid,
				html = ['<div z.c="', index,'" zs.t="STheader" class="zstopcell', zsw ? ' zsw' + zsw : '', '" ',
					zkS.t(zsw) ? 'z.zsw="' + zsw + '"' : '' ,'><div class="', 'zstopcelltxt ', zsw ? ' zswi' + zsw : '', '"',
					'>', name, '</div></div><div  class="zshboun"><div class="zshbouni" zs.t="SBoun"></div>'];
			if (hidden)
				html.push('<div class="zshbounw" zs.t="SBoun2"></div>');
			html.push('</div>');
			var $n = jq(html.join('')),
				cmp = $n[0],
				bcmp = $n[1];
			return new zss.Header(sheet, cmp, bcmp, index, type);
			
		} else if (type == zss.Header.VER) {
			//create row
			var html = ['<div z.r="', index, '" zs.t="SLheader" class="zsleftcell zsrow',
			            (zsh ? " zslh" + zsh : ''), '"', zkS.t(zsh) ? ' z.zsh="' + zsh + '"' : '', '><div class="zsleftcelltxt', 
			            (zsh ? " zslh" + zsh : ''), '">', name,'</div></div><div class="zsvboun"><div class="zsvbouni" zs.t="SBoun"></div>'];
			if (hidden)
				html.push('<div class="zsvbounw" zs.t="SBoun2"></div>');
			html.push('</div>');
			var $n = jq(html.join('')),
			    cmp = $n[0],
			    bcmp = $n[1];

			return new zss.Header(sheet, cmp, bcmp, index, type);
		}
	}
});
})();