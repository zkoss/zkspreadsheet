/* HeaderPositionHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 9, 2008 12:35:40 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import java.util.ArrayList;
import java.util.List;

/**
 * 
 * A utility class for calculating position of header
 * 
 * @author Dennis.Chen
 * 
 */
public class HeaderPositionHelper {

	int _defaultSize;
	int[][] _customizedSize;

	public HeaderPositionHelper(int defaultSize, int[][] customizedSize) {
		this._defaultSize = defaultSize;
		this._customizedSize = customizedSize;
	}

	public int[][] getCostomizedSize() {
		int[][] clone = new int[_customizedSize.length][2];
		System.arraycopy(_customizedSize, 0, clone, 0, _customizedSize.length);
		return clone;
	}

	public int getSize(int cellIndex) {
		// TODO binary search to speed up
		for (int i = 0; i < _customizedSize.length; i++) {
			if (cellIndex == _customizedSize[i][0])
				return _customizedSize[i][1];
			if (cellIndex < _customizedSize[i][0])
				return _defaultSize;
			continue;
		}
		return _defaultSize;
	}

	public void shiftMeta(int index, int size) {
		for (int i = 0; i < _customizedSize.length; i++) {
			if (index <= _customizedSize[i][0]) {
				_customizedSize[i][0] += size;
			}
		}
	}

	public void unshiftMeta(int index, int size) {
		List remove = new ArrayList();
		int end = index + size;
		for (int i = 0; i < _customizedSize.length; i++) {
			if (_customizedSize[i][0] >= index && _customizedSize[i][0] < end) {
				remove.add(Integer.valueOf(i));
			} else if (_customizedSize[i][0] >= end) {
				_customizedSize[i][0] -= size;
			}
		}
		end = remove.size();
		for (int i = end - 1; i >= 0; i--) {
			remove(((Integer) remove.get(i)).intValue());
		}
	}

	public int[] getMeta(int cellIndex) {
		for (int i = 0; i < _customizedSize.length; i++) {
			if (cellIndex == _customizedSize[i][0])
				return _customizedSize[i];
			if (cellIndex < _customizedSize[i][0])
				return null;
			continue;
		}
		return null;
	}

	public void setCustomizedSize(int cellIndex, int size, int id) {
		/*
		 * if(size==_defaultSize) { // don't remove, because there is a id
		 * attached. removeCustomizedSize(cellIndex); return; }
		 */
		int s = 0;
		int e = _customizedSize.length;
		if (e == 0) {
			insert(0, cellIndex, size, id);
			return;
		}
		int i;
		while (s < e) {// binary search
			i = s + (e - s) / 2;
			if (_customizedSize[i][0] == cellIndex) {
				_customizedSize[i][1] = size;
				_customizedSize[i][2] = id;
				return;
			} else if (_customizedSize[i][0] > cellIndex) {
				e = i - 1;
			} else {
				s = i + 1;
			}
			if (e == s) {
				if (e >= _customizedSize.length
						|| _customizedSize[e][0] > cellIndex) {
					insert(e, cellIndex, size, id);
				} else if (_customizedSize[e][0] == cellIndex) {
					_customizedSize[e][1] = size;
					_customizedSize[e][2] = id;
				} else {
					insert(e + 1, cellIndex, size, id);
				}
				break;
			} else if (e < s) {
				insert(s, cellIndex, size, id);
			}
		}
	}

	public void removeCustomizedSize(int cellIndex) {
		int s = 0;
		int e = _customizedSize.length;
		if (e == 0) {
			return;
		}
		int i;
		while (s < e) {// binary search
			i = s + (e - s) / 2;
			if (_customizedSize[i][0] == cellIndex) {
				remove(i);
				return;
			} else if (_customizedSize[i][0] > cellIndex) {
				e = i - 1;
			} else {
				s = i + 1;
			}
			if (e == s) {
				if (e < _customizedSize.length
						&& _customizedSize[e][0] == cellIndex) {
					remove(e);
				}
				break;
			} else if (e < s) {
				break;
			}
		}
	}

	private void insert(int index, int cellIndex, int size, int id) {
		int[][] result = new int[_customizedSize.length + 1][3];
		if (_customizedSize.length == 0) {
			_customizedSize = new int[][] { { cellIndex, size, id } };
			return;
		}
		if (index == 0) {
			result[0][0] = cellIndex;
			result[0][1] = size;
			result[0][2] = id;
			System.arraycopy(_customizedSize, 0, result, 1,
					_customizedSize.length);
		} else if (index > _customizedSize.length) {
			System.arraycopy(_customizedSize, 0, result, 0,
					_customizedSize.length);
			result[_customizedSize.length][0] = cellIndex;
			result[_customizedSize.length][1] = size;
			result[_customizedSize.length][2] = id;
		} else {
			System.arraycopy(_customizedSize, 0, result, 0, index);
			result[index][0] = cellIndex;
			result[index][1] = size;
			result[index][2] = id;
			System.arraycopy(_customizedSize, index, result, index + 1,
					_customizedSize.length - index);
		}
		_customizedSize = result;
	}

	private void remove(int index) {
		if (_customizedSize.length == 1) {
			_customizedSize = new int[0][];
			return;
		}
		int[][] result = new int[_customizedSize.length - 1][2];
		if (index == 0) {
			System.arraycopy(_customizedSize, 1, result, 0,
					_customizedSize.length - 1);
		} else if (index >= _customizedSize.length - 1) {
			System.arraycopy(_customizedSize, 0, result, 0,
					_customizedSize.length - 1);
		} else {
			System.arraycopy(_customizedSize, 0, result, 0, index);
			System.arraycopy(_customizedSize, index + 1, result, index,
					_customizedSize.length - index - 1);
		}
		_customizedSize = result;
	}

	public int getCellIndex(int px) {
		int sum = 0;
		int sum2 = 0;
		int index = 0;
		int inc = 0;
		if (px < 0)
			return 0;
		for (int i = 0; i < _customizedSize.length; i++) {
			sum2 = sum2 + (_customizedSize[i][0] - index) * _defaultSize;
			if (sum2 > px) {
				inc = px - sum;
				index = index + (inc / _defaultSize);// + (( inc % _defaultSize
														// == 0)?1:0);
				return index;
			}

			sum2 = sum2 + _customizedSize[i][1];
			if (sum2 > px) {
				return _customizedSize[i][0];
			}
			sum = sum2;
			index = _customizedSize[i][0] + 1;
		}

		inc = px - sum;
		index = index + (inc / _defaultSize);// + (( inc % _defaultSize ==
												// 0)?1:0);
		return index;
	}

	public int getStartPixel(int cellIndex) {
		int count = 0;
		int sum = 0;
		if (cellIndex < 0)
			return 0;
		for (int i = 0; i < _customizedSize.length; i++) {
			if (cellIndex <= _customizedSize[i][0]) {
				break;
			}
			count++;
			sum += _customizedSize[i][1];
		}
		sum = sum + (cellIndex - count) * _defaultSize;
		return sum;
	}

	public String toString() {
		StringBuffer sb = new StringBuffer();
		sb.append("[");
		for (int i = 0; i < _customizedSize.length; i++) {
			if (i != 0) {
				sb.append(",");
			}
			sb.append("[");
			sb.append(_customizedSize[i][0]).append(", ");
			sb.append(_customizedSize[i][1]).append(", ");
			sb.append(_customizedSize[i][2]);
			sb.append("]");

		}
		sb.append("]");
		return sb.toString();
	}

	public static void main(String args[]) {

		testGetCellIndex();
	}

	static public void testSetCustomizedSize() {
		int[][] customeizedSize = new int[0][];
		int defaultSize = 40;
		HeaderPositionHelper helper = new HeaderPositionHelper(defaultSize,
				customeizedSize);
		int id = 0;
		helper.setCustomizedSize(3, 30, id++);
		helper.setCustomizedSize(9, 20, id++);
		helper.setCustomizedSize(10, 60, id++);
		helper.setCustomizedSize(5, 60, id++);
		helper.setCustomizedSize(3, 60, id++);
		helper.setCustomizedSize(9, 30, id++);
		helper.setCustomizedSize(12, 60, id++);
		helper.setCustomizedSize(1, 60, id++);
		helper.setCustomizedSize(1, 20, id++);
		helper.setCustomizedSize(12, 90, id++);

		helper.setCustomizedSize(3, 40, id++);
		helper.setCustomizedSize(9, 40, id++);
		helper.setCustomizedSize(12, 40, id++);
		helper.setCustomizedSize(10, 40, id++);
		helper.setCustomizedSize(5, 40, id++);
		helper.setCustomizedSize(1, 40, id++);

		helper.setCustomizedSize(3, 30, id++);
		helper.setCustomizedSize(9, 20, id++);
		helper.setCustomizedSize(10, 60, id++);
		helper.setCustomizedSize(5, 60, id++);

		System.out.println("===>");

	}

	static public void testinsert() {
		int[][] customeizedSize = new int[0][];
		int defaultSize = 40;
		int id = 0;
		HeaderPositionHelper helper = new HeaderPositionHelper(defaultSize,
				customeizedSize);
		helper.insert(0, 5, 205, id++);
		helper.insert(0, 4, 204, id++);
		helper.insert(1, 6, 206, id++);
		helper.insert(2, 7, 207, id++);
		helper.insert(3, 8, 208, id++);
	}

	static public void testremove() {
		int[][] customeizedSize = new int[][] { { 4, 30 }, { 9, 20 },
				{ 10, 60 }, { 13, 3 }, { 14, 5 }, { 18, 8 }, { 23, 8 },
				{ 25, 8 } };
		int defaultSize = 40;
		HeaderPositionHelper helper = new HeaderPositionHelper(defaultSize,
				customeizedSize);
		helper.remove(customeizedSize.length);
		helper.remove(0);
		helper.remove(1);
		helper.remove(3);
		helper.remove(4);
		System.out.println("===>");
	}

	static public void testGetStartPixel() {
		int[][] customeizedSize = new int[][] { { 3, 30 }, { 9, 20 },
				{ 10, 60 } };
		int defaultSize = 40;
		HeaderPositionHelper helper = new HeaderPositionHelper(defaultSize,
				customeizedSize);
		System.out.println(">>>>" + helper.getStartPixel(0));// 0
		System.out.println(">>>>" + helper.getStartPixel(1));// 40
		System.out.println(">>>>" + helper.getStartPixel(3));// 120
		System.out.println(">>>>" + helper.getStartPixel(4));// 150
		System.out.println(">>>>" + helper.getStartPixel(5));// 190
		System.out.println(">>>>" + helper.getStartPixel(8));// 310
		System.out.println(">>>>" + helper.getStartPixel(9));// 350
		System.out.println(">>>>" + helper.getStartPixel(10));// 370
		System.out.println(">>>>" + helper.getStartPixel(11));// 430
	}

	static public void testGetCellIndex() {
		int[][] customeizedSize = new int[][] { { 3, 30 }, { 9, 20 },
				{ 10, 60 } };
		int defaultSize = 40;
		HeaderPositionHelper helper = new HeaderPositionHelper(defaultSize,
				customeizedSize);
		System.out.println(">>>>" + helper.getCellIndex(39));// 0
		System.out.println(">>>>" + helper.getCellIndex(40));// 1
		System.out.println(">>>>" + helper.getCellIndex(79));// 1
		System.out.println(">>>>" + helper.getCellIndex(80));// 2
		System.out.println(">>>>" + helper.getCellIndex(119));// 2
		System.out.println(">>>>" + helper.getCellIndex(120));// 3
		System.out.println(">>>>" + helper.getCellIndex(159));// 3
		System.out.println(">>>>" + helper.getCellIndex(160));// 4
		System.out.println(">>>>" + helper.getCellIndex(235));// 6
		System.out.println(">>>>" + helper.getCellIndex(369));// 9
		System.out.println(">>>>" + helper.getCellIndex(370));// 10
		System.out.println(">>>>" + helper.getCellIndex(429));// 10
		System.out.println(">>>>" + helper.getCellIndex(430));// 11
		System.out.println(">>>>" + helper.getCellIndex(480));// 12
	}

}
