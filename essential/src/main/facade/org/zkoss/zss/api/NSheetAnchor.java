package org.zkoss.zss.api;

/**
 * anchor of the object to attache to a sheet
 * @author dennis
 *
 */
public class NSheetAnchor {

	int row;
	int column;
	int lastRow;
	int lastColumn;
	int xOffset;
	int yOffset;
	int lastXOffset;
	int lastYOffset;

	public NSheetAnchor(int row, int column, int lastRow, int lastColumn) {
		this(row, column, 0, 0, lastRow, lastColumn, 0, 0);
	}

	public NSheetAnchor(int row, int column, int xOffset, int yOffset, int lastRow,
			int lastColumn, int lastXOffset, int lastYOffset) {
		this.row = row;
		this.column = column;
		this.xOffset = xOffset;
		this.yOffset = yOffset;
		this.lastRow = lastRow;
		this.lastColumn = lastColumn;
		this.lastXOffset = lastXOffset;
		this.lastYOffset = lastYOffset;
	}

	public int getRow() {
		return row;
	}

	public int getColumn() {
		return column;
	}

	public int getXOffset() {
		return xOffset;
	}

	public int getYOffset() {
		return yOffset;
	}

	public int getLastRow() {
		return lastRow;
	}

	public int getLastColumn() {
		return lastColumn;
	}

	public int getLastXOffset() {
		return lastXOffset;
	}

	public int getLastYOffset() {
		return lastYOffset;
	}

}
