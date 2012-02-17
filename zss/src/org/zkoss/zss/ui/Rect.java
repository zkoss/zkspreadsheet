/* Rect.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jun 4, 2008 9:32:24 AM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui;

/**
 * a class to represent a rectangle range with 4 value : left, top, right, bottom
 * @author Dennis.Chen
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
	
	public Object cloneSelf(){
		return (Rect)new Rect(_left,_top,_right,_bottom);
	}
	
	public String toString(){
		return "left:"+_left+",top:"+_top+",right:"+_right+",bottom:"+_bottom;
	}
	
	public boolean contains(int tRow, int lCol, int bRow, int rCol) {
		return	tRow >= _top && lCol >= _left &&
				bRow <= _bottom && rCol <= _right;
	}

	public int hashCode() {
		return _top << 14 + _left + _bottom << 14 + _right;
	}
	
	public boolean equals(Object obj){
		return (this == obj)
			|| (obj instanceof Rect 
					&& ((Rect)obj)._left == _left && ((Rect)obj)._right == _right 
					&& ((Rect)obj)._top == _top && ((Rect)obj)._bottom == _bottom);
	}
	
}
