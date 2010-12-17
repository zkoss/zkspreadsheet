/* Position.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jun 4, 2008 3:51:41 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui;

/**
 * a class to represent a cell position with 2 value : row,column
 * @author Dennis.Chen
 *
 */
public class Position {

	private int _row;
	private int _column;
	
	public Position(){
	}
	
	public Position(int row, int column){
		_row = row;
		_column = column;
	}
	
	public int getRow(){
		return _row;
	}
	
	/*package*/ void setRow(int row){
		_row = row;
	}
	
	public int getColumn(){
		return _column;
	}
	
	/*package*/ void setColumn(int column){
		_column = column;
	}
	
	@Override
	public String toString(){
		return "row:"+_row+",column:"+_column;
	}
	
	@Override
	public int hashCode() {
		return _row << 14 + _column;
	}
	
	@Override
	public boolean equals(Object obj){
		return (this == obj)
			|| (obj instanceof Position && ((Position)obj)._row == _row && ((Position)obj)._column == _column);
	}
}
