/* MainBlockCtrl.js

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
 * MainBlockCtrl handle scroll event and load spreadsheet content
 */
zss.MainBlockCtrl = zk.$extends(zss.CellBlockCtrl, {
	$init: function (sheet, cmp, left, top) {
		this.$supers('$init', arguments);
		this.loadstate = zss.MainBlockCtrl.IDLE;
	},
	/** load by server side block data, this method is invoke by server side command.**/
	_createCell: function (result) {
		var sheet = this.sheet,
			data = result.data,
			dir = result.dir,
			width = result.width,
			height = result.height,
			theaderdata = result.theader,
			lheaderdata = result.lheader,
			topfrozen = result.topfrozen,
			leftfrozen = result.leftfrozen;
		
		if (dir == "east") {
			this._createEastCell(data, width, height);
			if (theaderdata) {
				sheet.tp._createEastHeader(theaderdata, width);
			}
			
			if (topfrozen && sheet.tp.block) {
				var data = topfrozen.data,
					width = topfrozen.width,
					height = topfrozen.height;
				sheet.tp.block._createEastCell(data,width,height);
			}
		} else if (dir == "west") {
			this._createWestCell(data, width, height);
			if (theaderdata) sheet.tp._createWestHeader(theaderdata, width);
			if (topfrozen && sheet.tp.block) {
				data = topfrozen.data;
				width = topfrozen.width;
				height = topfrozen.height;
				sheet.tp.block._createWestCell(data, width, height);
			}
			this.sheet.dp._fixSize(this);
		} else if(dir == "south"){
			this._createSouthCell(data, width, height);
			if (lheaderdata) sheet.lp._createSouthHeader(lheaderdata, height);
			if (leftfrozen && sheet.lp.block) {
				data = leftfrozen.data;
				width = leftfrozen.width;
				height = leftfrozen.height;
				sheet.lp.block._createSouthCell(data, width, height);
			}
		} else if (dir == "north") {
			this._createNorthCell(data, width, height);
			if (lheaderdata) sheet.lp._createNorthHeader(lheaderdata, height);
			if (leftfrozen && sheet.lp.block) {
				data = leftfrozen.data;
				width = leftfrozen.width;
				height = leftfrozen.height;
				sheet.lp.block._createNorthCell(data, width, height);
			}
			this.sheet.dp._fixSize(this);
		} else if (dir == "jump") {
			
			this._createJumpCell(data, width, height);
			var tp = sheet.tp,
				lp = sheet.lp;
			if (theaderdata) tp._createJumpHeader(theaderdata, width);
			if (lheaderdata) lp._createJumpHeader(lheaderdata, height);
			
			if (topfrozen && tp.block) {

				data = topfrozen.data;
				width = topfrozen.width;
				height = topfrozen.height;
				var oldBlock = tp.block;
				tp.block = zss.CellBlockCtrl.createComp(sheet, result.left, 0, "zstopblock");
				tp.block._createJumpCell(data, width, height);
				
				tp.icomp.replaceChild(tp.block.comp, oldBlock.comp);
				oldBlock.cleanup();
			}
			if (leftfrozen && lp.block) {

				data = leftfrozen.data;
				width = leftfrozen.width;
				height = leftfrozen.height;
				var oldBlock = lp.block;
				lp.block = zss.CellBlockCtrl.createComp(sheet, 0, result.top, "zsleftblock");
				lp.block._createJumpCell(data, width, height);
				
				lp.icomp.replaceChild(lp.block.comp, oldBlock.comp);
				oldBlock.cleanup();
			}
		}
	},
	/**
	 * Handle scroll on spreadsheet.
	 */
	doScroll: function (vertical) {
		var sheet = this.sheet,
			range = zss.SSheetCtrl._getVisibleRange(sheet),
			alwaysjump = false;
		/*if(!sheet.snapedBlock && sheet.snapedBlock!=sheet.activeBlock){
			sheet._snapActiveBlock();
			alwaysjump = true; 
		}*/
		if (vertical) {
			var ctop = this.range.top,
				cbottom = this.range.bottom;
			if (range.top >= ctop && range.top <= cbottom) {
				if(range.bottom < cbottom)
					return; //the visible is be contained.
				//neighbor south
				sheet.activeBlock.loadCell(range.bottom, range.left, (alwaysjump ? -1 : 20), null, alwaysjump);
			} else if(ctop >= range.top && ctop <= range.bottom) {
				//neighbor north;
				sheet.activeBlock.loadCell(range.top,range.left, (alwaysjump ? -1 : 20), null, alwaysjump);
			} else if(range.top > cbottom) {
				//jump south
				sheet.activeBlock.loadCell(range.bottom, range.left,-1, null, alwaysjump);
			} else if(ctop > range.bottom) {
				//jump north;
				sheet.activeBlock.loadCell(range.top, range.left, -1, null, alwaysjump);
			} else{
				return;
			}
		} else {
			var cleft = this.range.left,
				cright = this.range.right;
			if (range.left >= cleft && range.left <= cright) {
				if (range.bottom < cbottom)
					return; //the visible is be contained.
				//neighbor east
				sheet.activeBlock.loadCell(range.top, range.right, (alwaysjump ? -1 : 5), null, alwaysjump);
			} else if (cleft >= range.left && cleft <= range.right) {
				//neighbor west;
				sheet.activeBlock.loadCell(range.top,range.left, (alwaysjump ? -1 : 5), null, alwaysjump);
			} else if (range.left>cright) {
				//jump east
				sheet.activeBlock.loadCell(range.top, range.right, -1, null, alwaysjump);
			} else if (cleft > range.right) {
				//jump west;
				sheet.activeBlock.loadCell(range.top, range.left, -1, null, alwaysjump);
			} else {
				return;
			}
		}
	},
	_sendOnCellFetch: function(token, type, direction, fetchLeft, fetchTop, fetchWidth, fetchHeight, vrange) {
		var sheet = this.sheet;
		this.loadstate = zss.MainBlockCtrl.LOADING;
		if (!vrange)
			vrange = zss.SSheetCtrl._getVisibleRange(sheet);

		var range = this.range,
			spcmp = this.sheet.sp.comp,
			dp = this.sheet.dp;

		sheet._wgt.fire('onZSSCellFetch', 
		 {token: token, sheetId: sheet.serverSheetId, type: type, direction: direction,
		 dpWidth: dp.width, dpHeight: dp.height, viewWidth: spcmp.clientWidth, viewHeight: spcmp.clientHeight,
		 blockLeft: range.left, blockTop: range.top, blockRight: range.right, blockBottom: range.bottom,
		 fetchLeft: fetchLeft, fetchTop: fetchTop, fetchWidth: fetchWidth, fetchHeight: fetchHeight,
		 rangeLeft: vrange.left, rangeTop: vrange.top, rangeRight: vrange.right, rangeBottom: vrange.bottom }, {toServer: true}, 25);
	},
	/**
	 * Returns Load Cell Direction
	 * @return string last direction (H: horizon; V: vertical; or null)
	 */
	_getLastDirection: function () {
		return this.sheet.activeBlock._lastDir;
	},
	/**
	 * load cell , if cell of row/col not exist, then load the cell into this block
	 * @param {Object} row row of a cell
	 * @param {Object} col col of a cell
	 * @param {Object} pruneres reserve when prune, -1 means don't prune, 0 means don't reserve 
	 * @param {Object} callbackfn callback function after cell be loaded
	 * @param {Object} alwaysjump always jump to this cell, this means skip this block, use a new bolck instead.
	 * @return true if already loaded, false if need to asynchronize loading.
	 */
	loadCell: function (row, col, pruneres, callbackfn, alwaysjump) {
		var cleft = this.range.left,
			ctop = this.range.top,
			cw = this.range.width,
			ch = this.range.height,
			cright = this.range.right,//cleft + cw - 1;
			cbottom = this.range.bottom;//ctop + ch - 1;
		
		if (row >= ctop && row <= cbottom && col >= cleft && col <= cright)
			return true;
		if (this.loadstate != zss.MainBlockCtrl.IDLE)//still waiting previous loading.
			return false;
		var token = "";
			sheet = this.sheet,
			range = zss.SSheetCtrl._getVisibleRange(sheet),
			local = this,
			fetchw = range.width, //size of cell of width to fetch back
			fetchh = range.height; //size of cell of height to fetch back
		
		if ((row >= ctop && row <= cbottom)) {//horizontal shift, east or west
			if (col < cleft) {//minus, west
				var y = ctop,
					h = ch,
					x = cleft - 1;
					w = x - col + 1;
				
				if (alwaysjump || w > fetchw) {
					//jump;
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "west", col, row, -1, -1, range);
				} else {
					//neighbor
					w = fetchw;
					if (x - w < 0)
						w = x + 1;

					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock,//get current activeBlock again.
							lastDir = block._lastDir; 
						block._lastDir = "H";
						if (lastDir && lastDir == "V"){
							block.pruneCell("south", range, 1);
							block.pruneCell("north", range, 1);
						} else if(pruneres >= 0) {
							block.pruneCell("east", range, pruneres);
						}
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if(callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "west", x, y, w, h, range);
				}

			} else {//positive, east
				var y = ctop,
					x = cleft + cw,
					w = col-x + 1;
				if (alwaysjump || w > fetchw) {
					//jump;
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "east", col, row, -1, -1, range);
				} else {
					//neighbor
					w = fetchw;
					var h = ch;
					if (x + w > this.sheet.maxCols)
						w = this.sheet.maxCols - x;
	
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock,//get current activeBlock again.
							lastDir = block._lastDir; 
						block._lastDir = "H";
						if (lastDir && lastDir == "V") {
							block.pruneCell("south", range, 1);
							block.pruneCell("north", range, 1);
						} else if (pruneres >= 0){
							block.pruneCell("west", range, pruneres);
						}
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "east", x, y, w, h, range);
				}
			}
		} else if ((col >= cleft && col <=cright)) {//vertical shift, shouth or north
			if (row < ctop) {//minus, north
				var x = cleft,
					w = cw,
					y = ctop - 1,
					h = y - row + 1;
				if (alwaysjump || h > fetchh) {
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "north", col, row, -1, -1, range);
				} else {
					h = fetchh;
					if (y - h < 0)
						h = y + 1;

					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock,//get current activeBlock again.
							lastDir = block._lastDir; 
						local._lastDir = "V";
						if (lastDir && lastDir == "H") {
							block.pruneCell("east", range, 1);
							block.pruneCell("west", range, 1);
						} else if (pruneres >= 0)
							block.pruneCell("south", range, pruneres);

						sheet.activeBlock.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();

						sheet.activeBlock.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "north", x, y, w, h, range);
				}
			} else {//positive, south
				var y = ctop + ch,
					x = cleft,
					w = cw,
					h = row-y +1 ;

				if (alwaysjump || h > fetchh) {
					//jump
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						block._lastDir = null;
						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "jump", "south", col, row, -1, -1, range);
				} else {
					h = fetchh;
					if (y + h > this.sheet.maxRows)
						h = this.sheet.maxRows - y;
					
					token = zkS.addCallback(function () {
						if (sheet.invalid) return;
						var block = sheet.activeBlock ;//get current activeBlock again.
						var lastDir = block._lastDir; 
						block._lastDir = "V";
						if (lastDir && lastDir == "H") {
							block.pruneCell("east", range, 1);
							block.pruneCell("west", range, 1);
						} else if(pruneres >= 0)
							block.pruneCell("north", range, pruneres);

						block.loadstate = zss.MainBlockCtrl.IDLE;
						if (callbackfn) callbackfn();
						block.loadForVisible();
					});
					this._sendOnCellFetch(token, "neighbor", "south", x, y, w, h, range);
				}
			}
		} else {
			token = zkS.addCallback(function () {
				if(sheet.invalid) return;
				var block = sheet.activeBlock ;//get current activeBlock again.
				block.loadstate = zss.MainBlockCtrl.IDLE;
				if(callbackfn) callbackfn();
				//local variable will be destoryed, don't use local.loadForVisible
				block.loadForVisible();
			});

			var direction = "";
			if (row < ctop && col < cleft) {
				//west-north
				direction = "westnorth";
			} else if(row < ctop && col > cright) {
				//east-north
				direction = "eastnorth";
			} else if(row > cbottom && col < cleft) {
				//west-south
				direction = "westsouth";
			} else if(row > cbottom && col > cright) {
				//east-shouth
				direction = "eastsouth";
			} else{
				direction = "eastnorth";
			}
			this._sendOnCellFetch(token, "jump", direction, col, row, -1, -1, range);
		}
		return false;
		
	},
	/**
	 * Prune cell
	 * @param string type the direction to prune
	 * @param zss.Range range the range to prune
	 * @param int reserve
	 */
	pruneCell: function (type , range, reserve) {
		var sheet = this.sheet;
		if (!sheet.config.prune) return;
		
		if (!range)
			range = zss.SSheetCtrl._getVisibleRange(sheet);
		if (!reserve || reserve < 0)
			reserve = 0;
		
		var sync = false;
		if (type == "west") {
			var l = range.left - reserve;
			l = sheet.mergeMatrix.getLeftConnectedColumn(l, this.range.top, this.range.bottom);

			if (l <= 0 || l < this.range.left) return;
			var size = l - this.range.left;
			if (size > 0) {
				this._removeWestCell(size);
				sheet.tp._removeWestHeader(size);
				if (sheet.tp.block)
					sheet.tp.block._removeWestCell(size);

				sync = true;
			}
		} else if (type == "east") {
			var r = range.right + reserve;
			r = sheet.mergeMatrix.getRightConnectedColumn(r,this.range.top,this.range.bottom);
			
			if(r > this.range.right) return;
			var size = this.range.right - r;
			if (size > 0) {
				this._removeEastCell(size);
				sheet.tp._removeEastHeader(size);
				if (sheet.tp.block)
					sheet.tp.block._removeEastCell(size);

				sync = true;
			}
		} else if (type == "north") {
			var t = range.top - reserve;
			t = sheet.mergeMatrix.getTopConnectedRow(t, this.range.left, this.range.right);
			
			if(t <= 0 || t < this.range.top) return;
			var size = t - this.range.top;
			if (size > 0) {
				this._removeNorthCell(size);
				sheet.lp._removeNorthHeader(size);
				if(sheet.lp.block)
					sheet.lp.block._removeNorthCell(size);

				sync = true;
			}
		} else if (type == "south") {
			var b = range.bottom + reserve;
			b = sheet.mergeMatrix.getBottomConnectedRow(b,this.range.left,this.range.right);
			
			if (b > this.range.bottom) return;
			var size = this.range.bottom - b;
			if (size > 0) {
				this._removeSouthCell(size);
				sheet.lp._removeSouthHeader(size);
				if(sheet.lp.block)
					sheet.lp.block._removeSouthCell(size);
			}
			sync = true;
		}
		if (sync) {
			//bug 1951423 IE : row is broken when scroll down
			//i must use timeout to delay processing
			setTimeout(function () {
				sheet.sendSyncblock();
			}, 0);	
		}
		sheet.dp._fixSize(this);
	},
	/**
	 * Load content and set the range to visible 
	 * @param zss.Range 
	 */
	loadForVisible: function (vrange) {
		var local = this,
			sheet = this.sheet;
		//TODO should i control this??
		//local.loadstate!=zkSMainBlockCtrl.IDLE
		if (!vrange)
			vrange = zss.SSheetCtrl._getVisibleRange(sheet);
		
		var top = this.range.top,
			left = this.range.left,
			right = this.range.right,
			bottom = this.range.bottom; 
		//visible range not cross this block,  this should invoke a jump 
		if (vrange.top > bottom || vrange.bottom < top || vrange.left > right || vrange.right < left) {
			//invoke a jump to top,left of visual range.
			this.loadCell(vrange.top, vrange.left, 0, null, true);

			return true;
		} else if (!(vrange.top >= top && vrange.left >= left && vrange.bottom <= bottom && vrange.right <= right)) {
			var token = zkS.addCallback(function () {
				//when inital , there is a loadForVisible, 
				//in this monent, use might click to invalidate this spreadhsheet
				//if click first, and then call loadForVisible.(those 2 event will send in one au_
				//then this call bak will throw exception because this block is invalidate.
				//TODO : so, in any call, i should check is it valid or not. 
				if (local.invalid) return;
				sheet.activeBlock.loadstate = zss.MainBlockCtrl.IDLE;
			});
			this._sendOnCellFetch(token, "visible", "", -1, -1, -1, -1, vrange);
			return true;
		}
		return false;
	},
	/**
	 * Reload block area content 
	 * @param string dir: "east", "west", "south", "north" or null
	 */
	reloadBlock: function (dir) {
		var local = this,
			sheet = this.sheet,
		//TODO should i control this??
		//local.loadstate!=zkSMainBlockCtrl.IDLE
			vrange = zss.SSheetCtrl._getVisibleRange(sheet),
			top = this.range.top,
			left = this.range.left,
			right = this.range.right,
			bottom = this.range.bottom; 
		var token = zkS.addCallback(function () {
			if (sheet.invald) return; 
			var block = sheet.activeBlock,
				range = block.range;
			block.loadstate = zss.MainBlockCtrl.IDLE;
			if (dir == "east" || dir == "west") {
				sheet.sp.scrollToVisible(null, dir == "east" ? range.right : range.left);
			} else if (dir == "south" || dir == "north") {
				sheet.sp.scrollToVisible(dir == "south" ? range.bottom: range.top ? range.right: range.left, null);
			}
			block.loadForVisible();
		});
		var col = dir=="east" ? right : left,
			row = dir=="south" ? bottom : top;
		this._sendOnCellFetch(token, "jump", (dir ? dir: "west"), col, row, -1, -1, vrange);
	}
}, {
	IDLE: 0,
	LOADING: 1,
	createComp: function (sheet, left, top, sclass) {	
		var cmp = document.createElement("div");
		jq(cmp).attr("zs.t", "SBlock").addClass(!sclass ? "zsblock" : sclass);
		return new zss.MainBlockCtrl(sheet, cmp, left, top);
	}
});
