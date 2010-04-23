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
	/**
	 * get cell index from a pixel
	 * @param int px
	 */
	getCellIndex: function (px) {
		if (px < 0) return 0;
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
			v1 = customizedSize[i][1];
			sum2 = sum2 + (v0 -index ) * defaultSize;
			if (sum2 > px) {
				inc = px - sum; 
				index = index + Math.floor(inc / defaultSize);
				return index;
			}
			sum2 = sum2 + v1;
			if (sum2 > px) {
				return v0;
			}
			sum = sum2;
			index = v0+1;
		}
		inc = px - sum; 
		index = index + Math.floor(inc / defaultSize);
		return index;
	},
	/**
	 * Returns start pixel(left or top) of a cell, 
	 * @param int cellIndex
	 * @return int
	 */
	getStartPixel: function (cellIndex) {
		if (cellIndex < 0) return 0;
		var count =0,
			sum = 0,
			customizedSize = this.custom,
			defaultSize = this.size,
			size = customizedSize.length;

		for (var i = 0; i < size; i++) {
			if (cellIndex <= customizedSize[i][0])
				break;

			count++;
			sum += customizedSize[i][1];
		}
		sum = sum + (cellIndex - count) * defaultSize;
		return sum;
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
	 */
	setCustomizedSize: function (cellIndex, size, id) {
		var customizedSize = this.custom,
			defaultSize = this.size;

		var s = 0,
			e = customizedSize.length;
		if (e == 0) {
			this._insert(0,cellIndex,size,id);
			return;
		}
		var i;
		while (s < e) {//binary search
			i = s + Math.floor((e - s) / 2);
			if (customizedSize[i][0] == cellIndex) {
				customizedSize[i][1] = size;
				customizedSize[i][2] = id;
				return;
			} else if (customizedSize[i][0] > cellIndex) {
				e = i - 1;
			} else {
				s = i + 1;
			}
			if (e == s) {
				if (e >= customizedSize.length || customizedSize[e][0] > cellIndex) {
					this._insert(e, cellIndex, size, id);
				} else if (customizedSize[e][0] == cellIndex) {
					customizedSize[e][1] = size;
					customizedSize[e][2] = id;
				} else {
					this._insert(e + 1, cellIndex, size, id);
				}
				break;
			} else if (e < s) {
				this._insert(s, cellIndex, size, id);
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
	_insert: function(index, cellIndex,size, id) {
		var customizedSize = this.custom;
		if (customizedSize.length == 0) {
			customizedSize.push.apply(customizedSize, [[cellIndex, size, id]]);
			return;
		}
		if (index == 0) {
			customizedSize.unshift([cellIndex, size, id]);
		} else if(index > customizedSize.length) {
			customizedSize.push.apply(customizedSize, [[cellIndex, size, id]]);
		} else {
			var tail = customizedSize.slice(index, customizedSize.length);
			customizedSize.length = index;
			customizedSize.push.apply(customizedSize, [[cellIndex, size, id]]);
			customizedSize.push.apply(customizedSize, tail);
		}
	}
});