package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.chart.*;

/**
 * Fill {@link NChartData} with series according to a selection. One series is determined by row or column. 
 * @author Hawk
 *
 */
//implementation references org.zkoss.zssex.api.ChartDataUtil
public class ChartDataHelper extends RangeHelperBase {

	public ChartDataHelper(NRange range) {
		super(range);
	}

	public void fillChartData(NChart chart) {
		NGeneralChartData chartData = (NGeneralChartData)chart.getData();
		CellRegion selection = new CellRegion(range.getRow(), range.getColumn(), range.getLastRow(), range.getLastColumn());
		switch (chart.getType()) {
		case AREA:
		case BAR:
		case COLUMN:
		case DOUGHNUT:
		case LINE:
		case PIE:
		case STOCK:
			fillCategoryData(selection, chartData);
			break;
		case SCATTER:
			fillXYData(selection, chartData);
			break;
		default:
			throw new UnsupportedOperationException();
		}
	}
	
	private void fillXYData(CellRegion selection, NGeneralChartData chartData) {
		CellRegion dataArea = getChartDataRange(selection);
		int firstDataColumn = dataArea.getColumn();
		int firstDataRow = dataArea.getRow();
		int columnCount = selection.getLastColumn() - firstDataColumn;
		int rowCount = selection.getLastRow() - firstDataRow;
		String xValueExpression = null;

		if (rowCount > columnCount) { 
			if (firstDataColumn < selection.getLastColumn()) {
				int lCol = selection.getColumn();
				int rCol = lCol;
				if (rCol < firstDataColumn) {
					rCol = firstDataColumn - 1;
				} else {
					firstDataColumn += 1;
				}
				//selection's first column becomes x values
				xValueExpression =new SheetRegion(sheet, new CellRegion(firstDataRow, lCol,selection.getLastRow(),rCol)).getReferenceString();
			}

			//each row is a series
			for (int seriesIndex = firstDataColumn ; seriesIndex <= dataArea.getLastColumn(); seriesIndex++){
				int nameRow = firstDataRow - 1;
				String nameExpression =null;
				if (nameRow >= selection.getRow()) {
					nameExpression = new SheetRegion(sheet, new CellRegion(selection.getRow(), seriesIndex, nameRow, seriesIndex) ).getReferenceString();
				}
				String yValueExpression =new SheetRegion(sheet, new CellRegion(dataArea.getRow(), seriesIndex, dataArea.getLastRow(), seriesIndex)).getReferenceString();
				NSeries series = chartData.addSeries();
				series.setXYFormula(nameExpression, xValueExpression, yValueExpression );
			}
		}else{
			if (firstDataRow < selection.getLastRow()) {
				int tRow = selection.getRow();
				int bRow = tRow;
				if (bRow < firstDataRow) {
					bRow = firstDataRow - 1;
				} else {
					firstDataRow += 1;
				}
				xValueExpression =new SheetRegion(sheet, new CellRegion(tRow,firstDataColumn, tRow,selection.getLastColumn())).getReferenceString();
			}

			//each column is a series
			for (int seriesIndex = firstDataRow ; seriesIndex <= dataArea.getLastRow(); seriesIndex++){
				String nameExpression = null;
				int nameCol = firstDataColumn - 1; //1 column to the left of data area is name column
				if (nameCol >= selection.getColumn()) {
					nameExpression = new SheetRegion(sheet, new CellRegion(seriesIndex, selection.getColumn(), seriesIndex, nameCol) ).getReferenceString();
				}
				String yValueExpression =new SheetRegion(sheet, new CellRegion(seriesIndex, dataArea.getColumn(), seriesIndex, dataArea.getLastColumn())).getReferenceString();
				NSeries series = chartData.addSeries();
				series.setXYFormula(nameExpression, xValueExpression, yValueExpression );
			}
		}
	}

	/**
	 * 
	 * @param sheet
	 * @param selection
	 * @param type
	 * @param anchor
	 * @return
	 */
	private void fillCategoryData(CellRegion selection, NGeneralChartData chartData) {
		CellRegion dataArea = getChartDataRange(selection);
		int firstDataColumn = dataArea.getColumn();
		int firstDataRow = dataArea.getRow();
		int columnCount = selection.getLastColumn() - firstDataColumn;
		int rowCount = selection.getLastRow() - firstDataRow;
		if (rowCount > columnCount) { //category by row, value by column
			int categoryColumn = firstDataColumn - 1; //1 column to the left of data area is category column
			if (categoryColumn >= selection.getColumn()) { //might not be present in selection area
				chartData.setCategoriesFormula(new SheetRegion(sheet, new CellRegion(dataArea.getRow(), categoryColumn, dataArea.getLastRow(), categoryColumn)).getReferenceString());
			}
			//each column is a series
			for (int seriesIndex = firstDataColumn ; seriesIndex <= dataArea.getLastColumn(); seriesIndex++){
				String nameExpression = null;
				int nameRow = firstDataRow - 1;
				if (nameRow >= selection.getRow()) {
					nameExpression = new SheetRegion(sheet, new CellRegion(selection.getRow(), seriesIndex, nameRow, seriesIndex) ).getReferenceString();
				}
				String xValueExpression =new SheetRegion(sheet, new CellRegion(dataArea.getRow(), seriesIndex, dataArea.getLastRow(), seriesIndex)).getReferenceString();
				NSeries series = chartData.addSeries();
				series.setFormula(nameExpression, xValueExpression );
			}
		}else{ //category by column, value by row
			int categoryRow = firstDataRow - 1; //1 row to the top of data area is category row
			if (categoryRow >= selection.getRow()) { 
				chartData.setCategoriesFormula(new SheetRegion(sheet, new CellRegion(categoryRow, dataArea.getColumn(), categoryRow, dataArea.getLastColumn())).getReferenceString());
			}
			//each row is a series
			for (int seriesIndex = firstDataRow ; seriesIndex <= dataArea.getLastRow(); seriesIndex++){
				String nameExpression = null;
				int nameColumn = firstDataColumn - 1;
				if (nameColumn >= selection.getColumn()) {
					nameExpression = new SheetRegion(sheet, new CellRegion(seriesIndex, selection.getColumn(), seriesIndex, nameColumn) ).getReferenceString();
				}
				String xValueExpression =new SheetRegion(sheet, new CellRegion(seriesIndex, dataArea.getColumn(), seriesIndex, dataArea.getLastColumn())).getReferenceString();
				NSeries series = chartData.addSeries();
				series.setFormula(nameExpression, xValueExpression );
			}
		}
	}
	
	/**
	 * Find a range of cells with only numeric value (or formula) inside selection. Skip those top and left side text headers.
	 * @param sheet
	 * @param selection
	 * @return
	 */
	private CellRegion getChartDataRange(CellRegion selection) {
		NSheet sheet = range.getSheet();
		// assume can't find number cell, use last cell as value
		int colIdx = selection.getColumn();
		int rowIdx = -1;
		for (int r = selection.getLastRow(); r >= selection.getRow(); r--) {
			NRow row = sheet.getRow(r);
			if(row==null) continue;
			int rCol = colIdx;
			for (int c = selection.getLastColumn(); c >= rCol; c--) {
				if (isQualifiedCell(sheet.getCell(r, c))) {
					colIdx = c;
					rowIdx = r;
				} else {
					break;
				}
			}
		}
		if (rowIdx == -1) { // can not find number cell, use last cell as chart's value
			rowIdx = selection.getLastRow();
			colIdx = selection.getLastColumn();
		}
		return new CellRegion(rowIdx, colIdx, selection.getLastRow(),
				selection.getLastColumn());
	}

	private boolean isQualifiedCell(NCell cell) {
		if (cell == null)
			return true;
		CellType cellType = cell.getType();
		//TODO formula might be evaluated to non-numeric
		return cellType == CellType.NUMBER
				|| cellType == CellType.FORMULA
				|| cellType == CellType.BLANK;
	}	
}
