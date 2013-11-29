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

}
