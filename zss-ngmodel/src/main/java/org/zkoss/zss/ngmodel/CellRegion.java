/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.util.AreaReference;
import org.zkoss.zss.ngmodel.util.CellReference;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class CellRegion implements Serializable {
	private static final long serialVersionUID = 1L;
	final public int row;
	final public int column;
	final public int lastRow;
	final public int lastColumn;
	AreaReference ref = null;

	public CellRegion(int row, int column) {
		this(row, column, row, column);
	}

	public CellRegion(String areaReference) {
		ref = new AreaReference(areaReference);
		this.row = ref.getFirstCell().getRow();
		this.column = ref.getFirstCell().getCol();
		this.lastRow = ref.getLastCell().getRow();
		this.lastColumn = ref.getLastCell().getCol();
		checkLegal();
	}
	
	public String getReferenceString(){
		if(ref==null){
			ref = new AreaReference(new CellReference(row,column),new CellReference(lastRow,lastColumn));
		}
		return ref.formatAsString();
	}

	private void checkLegal() {
		if ((row > lastRow && column > lastColumn)
				|| (row < 0 || lastRow < 0 || column < 0 || lastColumn < 0)) {
			throw new IllegalArgumentException("the region is legal " + this);
		}
	}

	public CellRegion(int row, int column, int lastRow, int lastColumn) {
		this.row = row;
		this.column = column;
		this.lastRow = lastRow;
		this.lastColumn = lastColumn;
		checkLegal();
	}

	public boolean isSingle() {
		return row == lastRow && column == lastColumn;
	}

	public boolean contains(int row, int column) {
		return row >= this.row && row <= this.lastRow && column >= this.column
				&& column <= this.lastColumn;
	}

	public boolean overlap(CellRegion region) {
		return contains(region.row, region.column)
				|| contains(region.row, region.lastColumn)
				|| contains(region.lastRow, region.column)
				|| contains(region.lastRow, region.lastColumn)
				|| region.contains(row, column);
	}

	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(row).append(",").append(column).append(",").append(lastRow)
				.append(",").append(lastColumn).append("]");

		return sb.toString();
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + column;
		result = prime * result + lastColumn;
		result = prime * result + lastRow;
		result = prime * result + row;
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		CellRegion other = (CellRegion) obj;
		if (column != other.column)
			return false;
		if (lastColumn != other.lastColumn)
			return false;
		if (lastRow != other.lastRow)
			return false;
		if (row != other.row)
			return false;
		return true;
	}
	
	

}
