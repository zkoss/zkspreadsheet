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

import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.poi.ss.util.CellReference;

/**
 * a class to represent a rectangle range with 4 value : left, top, right, bottom
 * @author Dennis.Chen
 *
 */
public class Rect {

//	public static final int SELECT_CELLS = 0x01;
//	public static final int SELECT_ROW = 0x02;
//	public static final int SELECT_COLUMN = 0x03;
//	public static final int SELECT_ALL = 0x04;
	
//	private int _action = SELECT_CELLS;//default
	private int _left = -1;
	private int _top = -1;
	private int _right = -1;
	private int _bottom = -1;
	
	private String areaRef;
	
	public Rect(){
	}
	
	public Rect(int left,int top,int right,int bottom){
		set(left,top,right,bottom);
	}
	public Rect(String areaReference){
		AreaReference ar = new AreaReference(areaReference);
		set(ar.getFirstCell().getCol(),ar.getFirstCell().getRow(),ar.getLastCell().getCol(),ar.getLastCell().getRow());
	}
	
//	public Rect(int action, int left,int top,int right,int bottom){
//		set(left,top,right,bottom);
//		_action = action;
//	}
	
//	public int getSelectionType(){
//		return _action;
//	}

	public void set(int left,int top,int right,int bottom){
		_left = left;
		_top = top;
		_right = right;
		_bottom = bottom;
		areaRef = null;
	}
	
	public int getLeft() {
		return _left;
	}

	public void setLeft(int left) {
		this._left = left;
		areaRef = null;
	}

	public int getTop() {
		return _top;
	}

	public void setTop(int top) {
		this._top = top;
		areaRef = null;
	}

	public int getRight() {
		return _right;
	}

	public void setRight(int right) {
		this._right = right;
		areaRef = null;
	}

	public int getBottom() {
		return _bottom;
	}

	public void setBottom(int bottom) {
		this._bottom = bottom;
		areaRef = null;
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
	
	public boolean overlap(int bTopRow, int bLeftCol, int bBottomRow, int bRightCol) {
		
		boolean xOverlap = isBetween(_left, bLeftCol, bRightCol) || isBetween(bLeftCol, _left, _right);
		boolean yOverlap = isBetween(_top, bTopRow, bBottomRow) || isBetween(bTopRow, _top, _bottom);
		
		return xOverlap && yOverlap;
	}
	
	private boolean isBetween(int value, int min, int max) {
		return (value >= min) && (value <= max);
	}
	
	/**
	 * @return Area reference of this selection
	 * @since 3.0.0
	 */
	public String toAreaReference(){
		if(areaRef==null){
			areaRef = new AreaReference(new CellReference(_top,_left),new CellReference(_bottom,_right)).formatAsString();
		}
		return areaRef;
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
	
	
	/**
	 * @return get first row, same as {@link #getTop()}
	 * @since 3.0.0
	 */
	public int getRow(){
		return _top;
	}
	
	/**
	 * @return get last row, same as {@link #getBottom()}
	 * @since 3.0.0
	 */
	public int getLastRow(){
		return _bottom;
	}
	
	/**
	 * @return get first column, same as {@link #getLeft()}
	 * @since 3.0.0
	 */
	public int getColumn(){
		return _left;
	}
	
	/**
	 * @return get last row, same as {@link #getRight()}
	 * @since 3.0.0
	 */
	public int getLastColumn(){
		return _right;
	}
}
