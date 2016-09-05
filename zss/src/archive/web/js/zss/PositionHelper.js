/* PositionHelper.js

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
	
	function updateSize(customizedSizeInfo, size, custom) {
		customizedSizeInfo[1] = size;
		customizedSizeInfo[4] = custom;
	}
/**
 * PositionHelper for calculate pixel to column/row index or vice versa
 */
zss.PositionHelper = zk.$extends(zk.Object, {
	$init: function (defaultSize, customizedSize) {
		this.$supers('$init', arguments);
		this.size = defaultSize;
		this.custom = customizedSize;
	},
	cleanup: function () {
		this.size = this.custom = null;
	},
	updateDefaultSize: function (defaultSize) {
		this.size = defaultSize;
	},
	getIncUnhidden: function (idx, limit) {
		var customizedSize = this.custom,
			size = customizedSize.length,
			v0,
			v1;
		for(var j = 0; j < size && idx <= limit; ++j) {
			v0 = customizedSize[j][0];
			if (idx == v0) {
				v1 = customizedSize[j][3]; //hidden: true
				if (!v1) {//special size but not hidden, return
					break;
				} else {
					++idx; //a hidden one, try next unhidden
				}
			} else if (idx < v0) { //normal case, return
				break; 
			}
		}
		return idx > limit ? -1 : idx;
	},
	getDecUnhidden: function (idx, limit) {
		var customizedSize = this.custom,
			size = customizedSize.length,
			v0,
			v1;
		for(var j = size - 1; j >= 0 && idx >= limit; --j) {
			v0 = customizedSize[j][0];
			if (idx == v0) {
				v1 = customizedSize[j][3]; //hidden: true
				if (!v1) {//special size but not hidden, return
					break;
				} else {
					--idx; //a hidden one, try next unhidden
				}
			} else if (idx > v0) { //normal case, return
				break; 
			}
		}
		return idx < limit ? -1 : idx;
	},
	/**
	 * get cell index from a pixel
	 * @param int px
	 */
	getCellIndex: function (px, firstHidden) { // ZSS-1266
		if (px < 0) px = 0; //col1/row1 can be hidden...
		var sum = 0,
			sum2 = 0,
			index = 0,
			inc = 0,
			customizedSize = this.custom,
			defaultSize = this.size,
			size = customizedSize.length,
			i = 0,
			v0,
			v1;
		for (;i < size; i++) {
			v0 = customizedSize[i][0];
			v1 = customizedSize[i][3] ? 0 : customizedSize[i][1]; //hidden, then size is zero
			sum2 = sum2 + (v0 -index ) * defaultSize;
			if (sum2 > px) {
				inc = px - sum;
				var incx = Math.floor(inc / defaultSize),
					cpx = inc - incx * defaultSize;
				index = index + incx;
				return [firstHidden ? this.getFirstHiddenIndex(i, index) : index, cpx]; //ZSS-1266
			}
			sum2 = sum2 + v1;
			if (sum2 > px) {
				var cpx = px - sum2 + v1;
				return [firstHidden ? this.getFirstHiddenIndex(i, v0) : v0, cpx]; //ZSS-1266
			}
			sum = sum2;
			index = v0+1;
		}
		inc = px - sum; 
		var incx = Math.floor(inc / defaultSize),
			cpx = inc - incx * defaultSize;
		index = index + incx;
		return [firstHidden ? this.getFirstHiddenIndex(i, index) : index, cpx]; //ZSS-1266
	},
	//ZSS-1266
	// find backward the 1st consecutive hidden row/col index 
	getFirstHiddenIndex: function (i, index) {
		while (--i >= 0 && this.custom[i][3] && this.custom[i][0] == index - 1) {
			// locate index
			--index;
		}
		return index;
	},
	/**
	 * Returns start pixel(left or top) of a cell, 
	 * @param int cellIndex
	 * @return int
	 */
	getStartPixel: function (cellIndex) {
		if (cellIndex < 0) return 0;
		return this._getStartPixel0(cellIndex)[0];
	},
	//ZSS-1117: return [start-pixel-of-cellIndex, new start metaIndex, new start pixel for the new start metaIndex]
	// + cellIndex to find its startPixel
	// + sidx - starting metaIndex; optional; default 0
	// + spx - starting pixel for the starting metaIndex; optional; default 0
	_getStartPixel0: function (cellIndex, sidx, spx) {
		var count = sidx || 0,
			sum = spx || 0,
			customizedSize = this.custom,
			defaultSize = this.size,
			size = customizedSize.length;

		for (var i = count; i < size; i++) {
			if (cellIndex <= customizedSize[i][0])
				break;

			count++;
			sum += customizedSize[i][3] ? 0 : customizedSize[i][1];
		}
		var spx0 = sum + (count > 0 ? (customizedSize[count-1][0] + 1 - count) : 0) * defaultSize; 
		sum = sum + (cellIndex - count) * defaultSize;
		return [sum, count, spx0];
	},
	//ZSS-1117
	getDiffPixel: function (startIdx, endIdx) {
		var sinfo = this._getStartPixel0(startIdx),
			einfo = this._getStartPixel0(endIdx+1, sinfo[1], sinfo[2]);
		return einfo[0] - sinfo[0];
	},
	//ZSS-1117
	getMetaIndex: function (cellIdx, min, max) {
		var customizedSize = this.custom,
			defaultSize = this.size,
			size = customizedSize.length,
			imin = min || 0,
			imax = max || size-1;
		while (imin <= imax) {
			var imid = Math.floor((imax + imin) / 2),
				cidx = customizedSize[imid][0];
			if ( cidx == cellIdx) {
				return imid;
			} else if (cidx < cellIdx) {
				imin = imid + 1; 
			} else {
				imax = imid - 1;
			}
		}
		return -imid - 1; // not found
	},
	/**
	 * Shift Meta 
	 * @param int cellIndex
	 * @param int size
	 */
	shiftMeta: function (cellIndex, size) {
		//TODO binary search to speed up
		var customizedSize = this.custom,
			len = customizedSize.length;
		for (var i = 0; i < len; i++) {
			if (cellIndex <= customizedSize[i][0] ) {
				customizedSize[i][0] += size;	
			}
		}
	},
	/**
	 * Unshift Meta
	 * @param int cellIndex
	 * @param int size
	 */
	unshiftMeta: function (cellIndex, size) {
		//TODO binary search to speed up
		var customizedSize = this.custom,
			len = customizedSize.length,
			remove = [],
			end = cellIndex + size,
			index;

		for (var i = 0; i < len; i++) {
			index = customizedSize[i][0];
			if (index >= cellIndex && index < end) {
				remove.push(i);
			} else if (index >= end) {
				customizedSize[i][0] -= size;	
			}
		}
		end = remove.length;
		for (var i = end - 1; i >= 0; i--) {
			this._remove(remove[i]);
		}
	},
	/**
	 * Returns whether cell's size is default or not
	 * 
	 * @param int index
	 * @return boolean indicate whether is default size or not
	 */
	isDefaultSize: function (index) {
		var defaultSize = this.size,
			size = this._getCustomizedSize(index),
			dif = Math.abs(size - defaultSize);
		if (Math.abs(size - defaultSize) > 1) { //default height tolerate range
			return false;
		}
		return true;
	},
	_getCustomizedSize: function (index) {
		var customizedSize = this.custom,
			i = customizedSize.length;
		while (i--) {
			if (index == customizedSize[i][0]) {
				return customizedSize[i][1];	
			}
		}
		return this.size;
	},
	/**
	 * Returns the size of the cell
	 * @param int cellIndex
	 * @return int size
	 */
	getSize: function (cellIndex) {
		//TODO binary search to speed up
		var customizedSize = this.custom,
			defaultSize = this.size,
			size = customizedSize.length;

		for (var i = 0; i < size; i++) {
			if (cellIndex == customizedSize[i][0]) return customizedSize[i][1];
			if (cellIndex < customizedSize[i][0]) return defaultSize;
		}
		return defaultSize;
	},
	/**
	 * Returns Meta
	 * @param int cellIndex
	 * @return array meta
	 */
	getMeta: function (cellIndex) {
		//TODO binary search to speed up
		var customizedSize = this.custom,
			defaultSize = this.size,
			size = customizedSize.length;

		for (var i = 0; i < size; i++) {
			if (cellIndex == customizedSize[i][0]) return customizedSize[i];
			if (cellIndex < customizedSize[i][0]) return null;
		}
		return null;
	},
	/**
	 * Sets the customized size of the cell
	 * @param int cellIndex
	 * @param int size
	 * @param string id
	 * @param boolean hidden
	 * @param boolean custom
	 */
	setCustomizedSize: function (cellIndex, size, id, hidden, custom) {
		var customizedSize = this.custom,
			defaultSize = this.size,
			s = 0,
			e = customizedSize.length;
		if (e == 0) {
			this._insert(0, cellIndex, size, id, hidden, custom);
			return;
		}
		var i;
		while (s < e) {//binary search
			i = s + Math.floor((e - s) / 2);
			if (customizedSize[i][0] == cellIndex) {
				customizedSize[i][2] = id;
				customizedSize[i][3] = hidden;
				updateSize(customizedSize[i], size, custom);
				return;
			} else if (customizedSize[i][0] > cellIndex) {
				e = i - 1;
			} else {
				s = i + 1;
			}
			if (e == s) {
				if (e >= customizedSize.length || customizedSize[e][0] > cellIndex) {
					this._insert(e, cellIndex, size, id, hidden, custom);
				} else if (customizedSize[e][0] == cellIndex) {
					customizedSize[e][2] = id;
					customizedSize[e][3] = hidden;
					updateSize(customizedSize[e], size, custom);
				} else {
					this._insert(e + 1, cellIndex, size, id, hidden, custom);
				}
				break;
			} else if (e < s) {
				this._insert(s, cellIndex, size, id, hidden, custom);
			}
		}
	},
	/**
	 * Remove the customized size of the cell
	 * @param int cellIndex
	 */
	removeCustomizedSize: function (cellIndex) {
		var customizedSize = this.custom,
			s = 0,
			e = customizedSize.length;
		if (e == 0)
			return;

		var i;
		while (s < e){//binary search
			i = s + Math.floor((e - s) / 2);
			if (customizedSize[i][0] == cellIndex) {
				this._remove(i);
				return;
			} else if (customizedSize[i][0] > cellIndex) {
				e = i - 1;
			} else {
				s = i + 1;
			}
			if (e == s) {
				if (e < customizedSize.length&&customizedSize[e][0] == cellIndex) {
					this._remove(e);
				}
				break;
			} else if (e < s) {
				break;
			}
		}
	},
	_remove: function (index) {
		var customizedSize = this.custom;
		if (customizedSize.length == 1) {
			customizedSize.pop();
			return;
		}
		if (index == 0) {
			customizedSize.shift();
		} else if (index >= customizedSize.length - 1) {
			customizedSize.pop();
		} else {
			var tail = customizedSize.slice(index + 1);
			customizedSize.length = index;
			customizedSize.push.apply(customizedSize, tail);
		}
	},
	_insert: function(index, cellIndex, size, id, hidden, custom) {
		var customizedSize = this.custom,
			obj = [cellIndex, size, id, hidden, custom];
		if (customizedSize.length == 0) {
			customizedSize.push.apply(customizedSize, [obj]);
			return;
		}
		if (index == 0) {
			customizedSize.unshift(obj);
		} else if(index > customizedSize.length) {
			customizedSize.push.apply(customizedSize, [obj]);
		} else {
			var tail = customizedSize.slice(index, customizedSize.length);
			customizedSize.length = index;
			customizedSize.push.apply(customizedSize, [obj]);
			customizedSize.push.apply(customizedSize, tail);
		}
	},
	/**
	 * An custom size is set manually users. The size determined by Spreadsheet automatically for some features, i.e. wrap text, 
	 * is not custom.
	 * @return boolean indicate whether the size is determined automatically or not
	 */
	isCustomSize: function(rowIndex){
		var customSize = this.getMeta(rowIndex);
		if (customSize == null){ //no custom info, default size
			return false;
		}else{
			return customSize[4];
		}
	},
	getDefaultSize: function(){
		return this.size;
	},
	/**
	 * @param rowColumnIndex row or column index
	 * @return return "hidden" value in position helper of specified row (or column).
	 * If no position helper exists, return FALSE.  
	 */
	isHidden: function(rowColumnIndex){
		var metaInfo =  this.getMeta(rowColumnIndex);
		if (metaInfo == null){
			return false;
		}else{
			return metaInfo[3];
		}
	}
});
})();