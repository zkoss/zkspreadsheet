/* AreaRefWithType.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2016/01/22, Created by Henri Chen	
}}IS_NOTE

Copyright (C) 2016 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.api;

import org.zkoss.zss.ui.CellSelectionType;

//ZSS-717
/**
 * A class that represents an area reference with 4 value : 
 * 		row(top row), column(left column), last row(bottom row) and last column(right column)
 *      and an extra selection type. (CELL, COL, ROW, ALL)

 * @author Henri Chen
 * @since 3.8.3
 */
public class AreaRefWithType extends AreaRef {
	private static final long serialVersionUID = -9010990229255383948L;
	
	protected CellSelectionType _type = CellSelectionType.CELL;
	
	public AreaRefWithType(int row,int column,int lastRow,int lastColumn, CellSelectionType type){
		setArea(row,column,lastRow,lastColumn);
		this._type = type;
	}
	public Object cloneSelf(){
		return (AreaRefWithType) new AreaRefWithType(_row,_column,_lastRow,_lastColumn, _type);
	}
	
	public boolean contains(int tRow, int lCol, int bRow, int rCol) {
		if (_type == CellSelectionType.ALL) { 
			return true;
		} else if (_type == CellSelectionType.COLUMN) {
			return tRow >= _row && bRow <= _lastRow;
		} else if (_type == CellSelectionType.ROW) {
			return	lCol >= _column && rCol <= _lastColumn;
		}
		return	tRow >= _row && lCol >= _column &&
				bRow <= _lastRow && rCol <= _lastColumn;
	}
	
	public boolean overlap(int bTopRow, int bLeftCol, int bBottomRow, int bRightCol) {
		boolean xOverlap = _type == CellSelectionType.COLUMN || isBetween(_column, bLeftCol, bRightCol) || isBetween(bLeftCol, _column, _lastColumn);
		boolean yOverlap = _type == CellSelectionType.ROW || isBetween(_row, bTopRow, bBottomRow) || isBetween(bTopRow, _row, _lastRow);

		return xOverlap && yOverlap;
	}

	public boolean overlap(AreaRefWithType areaRef) {
		return overlap(areaRef._row, areaRef._column, areaRef._lastRow, areaRef._lastColumn);
	}

	private boolean isBetween(int value, int min, int max) {
		return (value >= min) && (value <= max);
	}
	
	/**
	 * @return reference string, e.x A1:B2
	 */
	public String asString(){
		return new org.zkoss.poi.ss.util.AreaReference(new org.zkoss.poi.ss.util.CellReference(_row,_column),
					new org.zkoss.poi.ss.util.CellReference(_lastRow,_lastColumn)).formatAsString();
	}

	public int hashCode() {
		return _row << 14 + _column + _lastRow << 14 + _lastColumn + _type.ordinal() * 31;
	}
	
	public boolean equals(Object obj){
		return (this == obj)
			|| (obj instanceof AreaRefWithType 
					&& ((AreaRef)obj)._column == _column && ((AreaRef)obj)._lastColumn == _lastColumn 
					&& ((AreaRef)obj)._row == _row && ((AreaRef)obj)._lastRow == _lastRow
					&& ((AreaRefWithType)obj)._type == _type);
	}
	//ZSS-717
	/**
	 * 
	 * @param type
	 * @since 3.8.3
	 */
	public void setSelType(CellSelectionType type) {
		this._type = type;
	}
	
	//ZSS-717
	/**
	 * 
	 * @return
	 * @since 3.8.3
	 */
	public CellSelectionType getSelType() {
		return this._type;
	}

}
