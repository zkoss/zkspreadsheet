package org.zkoss.zss.api.model;

import java.util.ArrayList;

import org.zkoss.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.charts.CategoryData;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.usermodel.charts.ChartDataSource;
import org.zkoss.poi.ss.usermodel.charts.ChartTextSource;
import org.zkoss.poi.ss.usermodel.charts.DataSources;
import org.zkoss.poi.ss.usermodel.charts.XYData;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.charts.XSSFArea3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFAreaChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFBar3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFBarChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFColumn3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFColumnChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFDoughnutChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFLine3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFLineChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFPie3DChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFPieChartData;
import org.zkoss.poi.xssf.usermodel.charts.XSSFScatChartData;
import org.zkoss.zss.api.model.NChart.Type;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Rect;

/**
 * Utility to create simple ChartData from a sheet selection.
 * It guesses the series from the selection. 
 * @author dennis
 *
 */
//NOTE this class have to locate in EE
public class NChartDataUtil {

	public static NChartData getChartData(NSheet sheet, Rect selection,
			Type type) {
		ChartData data = newChartData(sheet, selection, type);
		return new NChartData(new SimpleRef<ChartData>(data));
	}

	// following code refers from SSWorkbookCtrl
	private static ChartData newChartData(NSheet sheet, Rect selection,
			Type type) {
		ChartData data = null;

		switch (type) {
		case Area3D:
			data = fillCategoryData(new XSSFArea3DChartData(), sheet, selection);
			break;
		case Area:
			data = fillCategoryData(new XSSFAreaChartData(), sheet, selection);
			break;
		case Bar3D:
			data = fillCategoryData(new XSSFBar3DChartData(), sheet, selection);
			// ((XSSFBar3DChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Column3D:
			data = fillCategoryData(new XSSFColumn3DChartData(), sheet, selection);
			// ((XSSFBar3DChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Bar:
			data = fillCategoryData(new XSSFBarChartData(), sheet, selection);
			// ((XSSFBarChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Column:
			data = fillCategoryData(new XSSFColumnChartData(), sheet, selection);
			// ((XSSFBarChartData) data).setGrouping(ChartGrouping.STANDARD);
			break;
		case Bubble:
			throw new UnsupportedOperationException();
		case Doughnut:
			data = fillCategoryData(new XSSFDoughnutChartData(), sheet, selection);
			break;
		case Line3D:
			data = fillCategoryData(new XSSFLine3DChartData(), sheet, selection);
			break;
		case Line:
			data = fillCategoryData(new XSSFLineChartData(), sheet, selection);
			break;
		case Pie3D:
			data = fillCategoryData(new XSSFPie3DChartData(), sheet, selection);
			break;
		case OfPie:
			// break;
			throw new UnsupportedOperationException();
		case Pie:
			data = fillCategoryData(new XSSFPieChartData(), sheet, selection);
			break;
		case Radar:
			throw new NotImplementedException("Radar data not impl");
		case Scatter:
			data = fillXYData(new XSSFScatChartData(), sheet, selection);
			break;
		case Stock:
			// data = fillCategoryData(new XSSFStockChartData());
			// break;
			throw new UnsupportedOperationException();
		case Surface3D:
			// break;
			throw new UnsupportedOperationException();
		case Surface:
			// break;
			throw new UnsupportedOperationException();
		}
		return data;
	}
	
	private static CellRangeAddress toCellRangeAddress(int row,int column,int lastRow,int lastColumn){
		return new CellRangeAddress(row,lastRow,column,lastColumn);
	}
	
	private static ChartData fillXYData(XYData data,NSheet sheet, Rect selection) {
		
		Rect rect = getChartDataRange(sheet,selection);
		int colIdx = rect.getLeft();
		int rowIdx = rect.getTop();
		
		ChartDataSource<Number> horValues = null;
		ArrayList<ChartTextSource> titles = new ArrayList<ChartTextSource>();
		ArrayList<ChartDataSource<Number>> values = new ArrayList<ChartDataSource<Number>>();
		
		int colWidth = selection.getRight() - colIdx;
		int rowHeight = selection.getBottom() - rowIdx;
		if (rowHeight > colWidth) {
			//find horizontal value, at least 1 column
			if (colIdx < selection.getRight()) {
				int lCol = selection.getLeft();
				int rCol = lCol;
				if (rCol < colIdx) {
					rCol = colIdx - 1;
				} else {
					colIdx += 1;
				}
//				String startCell = spreadsheet.getColumntitle(lCol) + spreadsheet.getRowtitle(rowIdx);
//				String endCell = spreadsheet.getColumntitle(rCol) + spreadsheet.getRowtitle(selection.getBottom());
				CellRangeAddress cra = toCellRangeAddress(rowIdx,lCol,selection.getBottom(),rCol);
				horValues = DataSources.fromNumericCellRange(sheet.getNative(), cra);
			}
			//find values
			int i = 1;
			for (int c = colIdx; c <= selection.getRight(); c++) {
				//find title
				String title = null;
				int row = rowIdx - 1;
				if (row >= selection.getTop()) {
					title = Ranges.range(sheet.getNative(), selection.getTop(), c, row, c).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
//				String startCell = spreadsheet.getColumntitle(c) + spreadsheet.getRowtitle(rowIdx);
//				String endCell = spreadsheet.getColumntitle(c) + spreadsheet.getRowtitle(selection.getBottom());
				CellRangeAddress cra = toCellRangeAddress(rowIdx,c,selection.getBottom(),c);
				values.add(DataSources.fromNumericCellRange(sheet.getNative(), cra));
			}
		} else {
			//find horizontal value, at least 1 row
			if (rowIdx < selection.getBottom()) {
				int tRow = selection.getTop();
				int bRow = tRow;
				if (bRow < rowIdx) {
					bRow = rowIdx - 1;
				} else {
					rowIdx += 1;
				}
//				String startCell = spreadsheet.getColumntitle(colIdx) + spreadsheet.getRowtitle(tRow);
//				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(tRow);
				CellRangeAddress cra = toCellRangeAddress(tRow,colIdx,tRow,selection.getRight());
				horValues = DataSources.fromNumericCellRange(sheet.getNative(), cra);
			}
			//find values
			int i = 1;
			for (int r = rowIdx; r <= selection.getBottom(); r++) {
				//find title
				String title = null;
				int col = colIdx - 1;
				if (col >= selection.getLeft()) {
					title = Ranges.range(sheet.getNative(), r, selection.getLeft(), r, col).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));
				
//				String startCell = spreadsheet.getColumntitle(colIdx) + spreadsheet.getRowtitle(r);
//				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(r);
				CellRangeAddress cra = toCellRangeAddress(r,colIdx,r,selection.getRight());
				values.add(DataSources.fromNumericCellRange(sheet.getNative(), cra));
			}
		}
		
		for (int i = 0; i < values.size(); i++) {
			data.addSerie(titles.get(i), horValues, values.get(i));
		}
		return data;
	}

	private static CategoryData fillCategoryData(CategoryData data,
			NSheet sheet, Rect selection) {
		Rect rect = getChartDataRange(sheet, selection);
		int colIdx = rect.getLeft();
		int rowIdx = rect.getTop();

		ChartDataSource<String> cats = null;
		ArrayList<ChartTextSource> titles = new ArrayList<ChartTextSource>();
		ArrayList<ChartDataSource<Number>> vals = new ArrayList<ChartDataSource<Number>>();

		int colWidth = selection.getRight() - colIdx;
		int rowHeight = selection.getBottom() - rowIdx;
		if (rowHeight > colWidth) { // catalog by row, value by column
			// find catalog
			int col = colIdx - 1;
			if (col >= selection.getLeft()) {
//				String startCell = spreadsheet.getColumntitle(selection.getLeft()) + spreadsheet.getRowtitle(rowIdx);
//				String endCell = spreadsheet.getColumntitle(col)+ spreadsheet.getRowtitle(selection.getBottom());
				CellRangeAddress cra = toCellRangeAddress(rowIdx,selection.getLeft(),selection.getBottom(),col);
				cats = DataSources.fromStringCellRange(sheet.getNative(), cra);
			}
			// find value, by column
			int i = 1;
			for (int c = colIdx; c <= selection.getRight(); c++) {
				// find title
				String title = null;
				int row = rowIdx - 1;
				if (row >= selection.getTop()) {
					title = Ranges.range(sheet.getNative(), selection.getTop(), c, row, c).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));

//				String startCell = spreadsheet.getColumntitle(c)+ spreadsheet.getRowtitle(rowIdx);
//				String endCell = spreadsheet.getColumntitle(c)+ spreadsheet.getRowtitle(selection.getBottom());
				CellRangeAddress cra = toCellRangeAddress(rowIdx,c,selection.getBottom(),c);
				ChartDataSource<Number> val = DataSources.fromNumericCellRange(sheet.getNative(),cra);
				vals.add(val);
			}
		} else { // catalog by column, value by row
			// find catalog
			int row = rowIdx - 1;
			if (row >= selection.getTop()) {
//				String startCell = spreadsheet.getColumntitle(colIdx)+ spreadsheet.getRowtitle(row);
//				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(row);
				CellRangeAddress cra = toCellRangeAddress(row,colIdx,row,selection.getRight());
				cats = DataSources.fromStringCellRange(sheet.getNative(),cra);
			}

			// find value
			int i = 1;
			for (int r = rowIdx; r <= selection.getBottom(); r++) {
				// find title
				String title = null;
				int col = colIdx - 1;
				if (col >= selection.getLeft()) {
					title = Ranges.range(sheet.getNative(), r, selection.getLeft(), r,
									col).getText().toString();
				}
				titles.add(title == null ? null : DataSources.fromString(title));

//				String startCell = spreadsheet.getColumntitle(colIdx)+ spreadsheet.getRowtitle(r);
//				String endCell = spreadsheet.getColumntitle(selection.getRight()) + spreadsheet.getRowtitle(r);
				CellRangeAddress cra = toCellRangeAddress(r,colIdx,r,selection.getRight());
				ChartDataSource<Number> val = DataSources.fromNumericCellRange(sheet.getNative(),cra);
				vals.add(val);
			}
		}

		for (int i = 0; i < vals.size(); i++) {
			data.addSerie(titles.get(i), cats, vals.get(i));
		}
		return data;
	}

	private static Rect getChartDataRange(NSheet sheet,Rect selection) {
		// assume can't find number cell, use last cell as value
		int colIdx = selection.getLeft();
		int rowIdx = -1;
		Worksheet ws = sheet.getNative();
		for (int r = selection.getBottom(); r >= selection.getTop(); r--) {
			Row row = ws.getRow(r);
			int rCol = colIdx;
			for (int c = selection.getRight(); c >= rCol; c--) {
				if (isQualifiedCell(row.getCell(c))) {
					colIdx = c;
					rowIdx = r;
				} else {
					break;
				}
			}
		}
		if (rowIdx == -1) { // can not find number cell, use last cell as
							// chart's value
			rowIdx = selection.getBottom();
			colIdx = selection.getRight();
		}
		return new Rect(colIdx, rowIdx, selection.getRight(),
				selection.getBottom());
	}

	private static boolean isQualifiedCell(Cell cell) {
		if (cell == null)
			return true;
		int cellType = cell.getCellType();
		return cellType == Cell.CELL_TYPE_NUMERIC
				|| cellType == Cell.CELL_TYPE_FORMULA
				|| cellType == Cell.CELL_TYPE_BLANK;
	}

}
