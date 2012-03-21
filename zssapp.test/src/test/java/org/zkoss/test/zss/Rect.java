/* Rect.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 8:50:48 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import com.google.common.base.Objects;

/**
 * @author sam
 *
 */
public class Rect {
	
	private int _left = -1;
	private int _top = -1;
	private int _right = -1;
	private int _bottom = -1;
	
	public Rect(){
	}
	
	public Rect(int left,int top,int right,int bottom){
		set(left,top,right,bottom);
	}

	public void set(int left,int top,int right,int bottom){
		_left = left;
		_top = top;
		_right = right;
		_bottom = bottom;
	}
	
	public int getLeft() {
		return _left;
	}

	public void setLeft(int left) {
		this._left = left;
	}

	public int getTop() {
		return _top;
	}

	public void setTop(int top) {
		this._top = top;
	}

	public int getRight() {
		return _right;
	}

	public void setRight(int right) {
		this._right = right;
	}

	public int getBottom() {
		return _bottom;
	}

	public void setBottom(int bottom) {
		this._bottom = bottom;
	}

	@Override
	public boolean equals(Object obj) {
		if (obj instanceof Rect) {
			Rect that = (Rect) obj;
			return Objects.equal(this._left, that._left)
				&& Objects.equal(this._top, that._top)
				&& Objects.equal(this._right, that._right)
				&& Objects.equal(this._bottom, that._bottom);
		}
		return false;
	}

	@Override
	public String toString() {
		return Objects.toStringHelper(this)
			.add("top", _top)
			.add("left", _left)
			.add("bottom", _bottom)
			.add("right", _right)
			.toString();
	}
	
	
}
