/* SheetCtrlImpl.java

	Purpose:
		
	Description:
		
	History:
		Dec 14, 2010 3:20:42 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.Combo;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.PivotCache;
import org.zkoss.poi.ss.usermodel.PivotTable;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.ZssChartX;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.usermodel.charts.ChartGrouping;
import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.poi.ss.usermodel.charts.LegendPosition;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;

/**
 * Common implementation of the {@link SheetCtrl} interface. 
 * @author henrichen
 *
 */
public class SheetCtrlImpl implements SheetCtrl {
	private final Book _book;
	protected final Worksheet _sheet;
	private boolean _evalAll;
	private String _uuid;
	
	public SheetCtrlImpl(Book book, Worksheet sheet) {
		_book = book;
		_sheet = sheet;
	}
    /*package*/ void initUuid() {
    	if (_uuid == null) {
    		_uuid = (String) ((BookCtrl)_book).nextSheetId();
    	}
    }
    
	@Override
	public void evalAll() {
		// TODO Auto-generated method stub
		for(Row row : _sheet) {
			if (row != null) {
				for(Cell cell : row) {
					if (cell != null && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						BookHelper.evaluate(_sheet.getBook(), cell);
					}
				}
			}
		}
		_evalAll = true;
	}

	@Override
	public boolean isEvalAll() {
		return _evalAll;
	}
	
	@Override
	public String getUuid() {
    	if (_uuid == null) {
    		_uuid = (String)((BookCtrl)_book).nextSheetId();
    	}
		return _uuid;
	}
    private Map<String, CellRangeAddress> _mergedRegions;
	private Map<String, CellRangeAddress> getMergedRegions() { //ZSS-77: Drag fill "Jan" on new sheet, cause NPE
		if (_mergedRegions == null)
			initMerged();
		return _mergedRegions;
	}
	@Override
	public void initMerged() {
    	final int num = _sheet.getNumMergedRegions();
    	_mergedRegions = new HashMap<String, CellRangeAddress>(num);
    	for (int j = 0; j < num; ++j) {
    		final CellRangeAddress addr = _sheet.getMergedRegion(j);
    		final int col = addr.getFirstColumn();
    		final int row = addr.getFirstRow();
    		_mergedRegions.put(mergeId(row, col), addr);
    	}
	}
	@Override
	public CellRangeAddress getMerged(int row, int col) {
		final Map<String, CellRangeAddress> mergedRegions = getMergedRegions(); 
		return mergedRegions.get(mergeId(row, col));	
	}
	@Override
	public void addMerged(CellRangeAddress addr) {
		final int tRow = addr.getFirstRow();
		final int lCol = addr.getFirstColumn();
		final Map<String, CellRangeAddress> mergedRegions = getMergedRegions();
		mergedRegions.put(mergeId(tRow, lCol), addr);
	}
	@Override
	public void deleteMerged(CellRangeAddress addr) {
		final int tRow = addr.getFirstRow();
		final int lCol = addr.getFirstColumn();
		final Map<String, CellRangeAddress> mergedRegions = getMergedRegions();
		mergedRegions.remove(mergeId(tRow, lCol));
	}
	private String mergeId(int row, int col) {
		return row+"_"+col;
	}
	
	@Override
	public void whenRenameSheet(String oldname, String newname) {
		//do nothing.
	}
	
	@Override
	public DrawingManager getDrawingManager() {
		return new DrawingManager() {
			@Override
			public List<ZssChartX> getChartXs() {
				return Collections.emptyList();
			}

			@Override
			public List<Picture> getPictures() {
				return Collections.emptyList();
			}

			@Override
			public List<Combo> getCombos() {
				return Collections.emptyList();
			}

			@Override
			public ZssChartX addChartX(Worksheet sheet, ClientAnchor anchor,
					ChartData data, ChartType type, ChartGrouping grouping,
					LegendPosition pos) {
				return null;
			}

			@Override
			public Picture addPicture(Worksheet sheet, ClientAnchor anchor,
					byte[] imageData, int format) {
				return null;
			}

			@Override
			public void deletePicture(Worksheet sheet, Picture picture) {
			}

			@Override
			public void movePicture(Worksheet sheet, Picture picture,
					ClientAnchor anchor) {
			}

			@Override
			public void moveChart(Worksheet sheet, Chart chart,
					ClientAnchor anchor) {
			}

			@Override
			public List<Chart> getCharts() {
				return null;
			}

			@Override
			public void deleteChart(Worksheet sheet, Chart chart) {
			}
		};
	}
}
