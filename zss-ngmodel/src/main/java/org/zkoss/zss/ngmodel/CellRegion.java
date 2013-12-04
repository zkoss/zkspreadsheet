package org.zkoss.zss.ngmodel;

import java.io.Serializable;

public class CellRegion implements Serializable {
	private static final long serialVersionUID = 1L;
	final public int row;
	final public int column;
	final public int lastRow;
	final public int lastColumn;

	public CellRegion(int row, int column) {
		this(row, column, row, column);
	}

	public CellRegion(int row, int column, int lastRow, int lastColumn) {
		this.row = row;
		this.column = column;
		this.lastRow = lastRow;
		this.lastColumn = lastColumn;
	}
	
	public boolean isSingle(){
		return row == lastRow && column == lastColumn;
	}

	
	public boolean contains(int row, int column){
		return row>=this.row && row<=this.lastRow && column>=this.column && column<=this.lastColumn;
	}
	
	public boolean contains(CellRegion region){
		return contains(region.row,region.column) ||
				contains(region.row,region.lastColumn) ||
				contains(region.lastRow,region.column) ||
				contains(region.lastRow,region.lastColumn) ||
				region.contains(row,column);
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("[").append(row).append(column).append(lastRow).append(lastColumn).append("]");
		
		return sb.toString();
	}
	
}
