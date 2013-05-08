package org.zkoss.zss.api;

import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Picture;

/**
 * The anchor for the objects ( {@link Picture} , {@link Chart}) to attache to a sheet
 * @author dennis
 * @see Picture
 * @see Range#addPicture(SheetAnchor, byte[], org.zkoss.zss.api.model.Picture.Format)
 * @see Chart
 * @see Range#addChart(SheetAnchor, org.zkoss.zss.api.model.ChartData, org.zkoss.zss.api.model.Chart.Type, org.zkoss.zss.api.model.Chart.Grouping, org.zkoss.zss.api.model.Chart.LegendPosition)
 */
public class SheetAnchor {

	int row;
	int column;
	int lastRow;
	int lastColumn;
	int xOffset;
	int yOffset;
	int lastXOffset;
	int lastYOffset;

	public SheetAnchor(int row, int column, int lastRow, int lastColumn) {
		this(row, column, 0, 0, lastRow, lastColumn, 0, 0);
	}

	public SheetAnchor(int row, int column, int xOffset, int yOffset, int lastRow,
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
