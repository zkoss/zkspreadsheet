/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.*;
import java.util.*;

import org.zkoss.poi.hssf.model.HSSFFormulaParser;
import org.zkoss.poi.hssf.record.chart.LinkedDataRecord;
import org.zkoss.poi.hssf.usermodel.*;
import org.zkoss.poi.hssf.usermodel.HSSFChart.HSSFChartType;
import org.zkoss.poi.hssf.usermodel.HSSFChart.HSSFSeries;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.chart.*;

/**
 * 
 * @author Hawk
 * @since 3.5.0
 */
public class NExcelXlsImporter extends AbstractExcelImporter{

	
	@Override
	protected Workbook createPoiBook(InputStream is) throws IOException{
		return new HSSFWorkbook(is);
	}

	@Override
	protected void importExternalBookLinks() {
		// do nothing
		// xls file has no individual external book links
		// every reference has every necessary information including external book index
		// and already resolved by POI
	}
	
	/**
	 * 
	 * @param poiSheet
	 * @return 256
	 */
	private int getLastChangedColumnIndex(Sheet poiSheet) {
		return new HSSFSheetHelper((HSSFSheet)poiSheet).getInternalSheet().getMaxConfiguredColumn();
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void importColumn(Sheet poiSheet, NSheet sheet, int defaultWidth) {
		int lastChangedColumnIndex = getLastChangedColumnIndex(poiSheet);
		for (int c=0 ; c <= lastChangedColumnIndex ; c++){
			//reference Spreadsheet.updateColWidth()
			NColumn col = sheet.getColumn(c);
			int width = XUtils.getWidthAny(poiSheet, c, CHRACTER_WIDTH);
			boolean hidden = poiSheet.isColumnHidden(c);
			col.setHidden(hidden);
			//to avoid creating unnecessary column with just default value
			if(!hidden && width != defaultWidth){
				col.setWidth(width);
			}
			CellStyle columnStyle = poiSheet.getColumnStyle(c); 
			if (columnStyle != null){
				col.setCellStyle(importCellStyle(columnStyle));
			}
		}
	}

	@Override
	protected void importChart(Sheet poiSheet, NSheet sheet) {
		List<ZssChartX> charts = importHSSFDrawings((HSSFSheet)poiSheet);
		//reference ChartHelper.drawHSSFChart()
		for (ZssChartX zssChart : charts){
			NViewAnchor viewAnchor = toViewAnchor(poiSheet, zssChart.getPreferredSize());
			final HSSFChart hssfChart = (HSSFChart)zssChart.getChartInfo();
			HSSFChartType type = hssfChart.getType();
			NChart chart = null;
			if (type != null){
				switch(type) {
				case Area:
					chart = sheet.addChart(NChartType.AREA, viewAnchor);
					break;
				case Bar:
					//				model = prepareBarModel(drawer, (HSSFSheet)sheet, series);
//					break;
				case Line:
					//				model = prepareLineModel(drawer, (HSSFSheet)sheet, series);
//					break;
				case Pie:
					//				model = preparePieModel(drawer, (HSSFSheet)sheet, series);
//					break;
				case Scatter:
					//				model = prepareScatterModel(drawer, (HSSFSheet)sheet, series); //ZSS-107
//					break;
				default:
					continue;
				}
			}
			importSeries(Arrays.asList(hssfChart.getSeries()), (NGeneralChartData)chart.getData());
			if (hssfChart.getChartTitle() != null){
				chart.setTitle(hssfChart.getChartTitle());
			}
		}

	}

	/**
	 * reference DrawingManagerImpl.initHSSFDrawings()
	 * @param sheet
	 * @return
	 */
	private List<ZssChartX> importHSSFDrawings(HSSFSheet sheet) {
		List<ZssChartX> charts = new LinkedList<ZssChartX>();
		//decode drawing/obj/chartStream record into shapes and construct the shape tree in patriarch
		final HSSFPatriarch patriarch = sheet.getDrawingPatriarch(); 
		//will call sheet.getDrawingEscherAggregate() 
		//and try to convert Record to HSSFShapes but will NOT!
		if (patriarch != null) {
			HSSFPatriarchHelper helper = new HSSFPatriarchHelper(patriarch);

			//populate pictures and charts
			for (HSSFShape shape : patriarch.getChildren()) {
				if (shape instanceof HSSFChartShape) {
					new HSSFChartDecoder(helper,(HSSFChartShape)shape).decode();
					charts.add((HSSFChartShape)shape);
				} else {
					//log "unprocessed shape"
				}
			}
		}
		return charts;
	}

	/**
	 * reference DefaultBookWidgetLoader.getHSSFHeightInPx()
	 * @param anchor
	 * @param poiSheet
	 * @return
	 */
	@Override
	protected int getAnchorHeightInPx(ClientAnchor anchor, Sheet poiSheet) {
	    final int t = anchor.getRow1();
	    final int tfrc = anchor.getDy1();
	    
	    //first row
	    final int th = XUtils.getHeightAny(poiSheet,t);
	    final int hFirst = tfrc >= 256 ? 0 : (th - (int) Math.round(((double)th) * tfrc / 256));  
	    
	    //last row
	    final int b = anchor.getRow2();
	    int hLast = 0;
	    if (t != b) {
		    final int bfrc = anchor.getDy2();
	    	final int bh = XUtils.getHeightAny(poiSheet,b);
	    	hLast = (int) Math.round(((double)bh) * bfrc / 256);  
	    }
	    
	    //in between
	    int height = hFirst + hLast;
	    for (int j = t+1; j < b; ++j) {
	    	height += XUtils.getHeightAny(poiSheet,j);
	    }
	    
	    return height;
	}
	
	/**
	 * reference DefaultBookWidgetLoader.getHSSFWidthInPx()
	 * @param anchor
	 * @param sheet
	 * @return
	 */
	@Override
	protected int getAnchorWidthInPx(ClientAnchor anchor, Sheet sheet) {
	    final int firstColumn = anchor.getCol1();
	    final int lfrc = anchor.getDx1();
	    
	    //first column
	    final int lw = XUtils.getWidthAny(sheet,firstColumn, CHRACTER_WIDTH);
	    
	    final int wFirst = lfrc >= 1024 ? 0 : (lw - (int) Math.round(((double)lw) * lfrc / 1024));  
	    
	    //last column
	    final int lastColumn = anchor.getCol2();
	    int wLast = 0;
	    if (firstColumn != lastColumn) {
		    final int rfrc = anchor.getDx2();
	    	final int rw = XUtils.getWidthAny(sheet,lastColumn, CHRACTER_WIDTH);
	    	wLast = (int) Math.round(((double)rw ) * rfrc / 1024);  
	    }
	    
	    //in between
	    int width = wFirst + wLast;
	    for (int col = firstColumn+1; col < lastColumn; ++col) {
	    	width += XUtils.getWidthAny(sheet,col, CHRACTER_WIDTH);
	    }

	    return width;
	}

	private void importSeries(List<HSSFSeries> seriesList, NGeneralChartData chartData) {
		HSSFSeries firstSeries = null;
		if ((firstSeries = seriesList.get(0))!=null){
			chartData.setCategoriesFormula(getTitleFormula(firstSeries.getDataCategoryLabels()));
		}
		for (int i =0 ;  i< seriesList.size() ; i++){
			HSSFSeries sourceSeries = seriesList.get(i);
			String nameExpression = getTitleFormula(sourceSeries.getDataName());			
			String xValueExpression = getValueFormula(sourceSeries.getDataValues());
			NSeries series = chartData.addSeries();
			series.setFormula(nameExpression, xValueExpression);
		}
	}
	

	private String getValueFormula(LinkedDataRecord dataValues) {
		if (dataValues == null) {
			return "0";
		}
		return HSSFFormulaParser.toFormulaString((HSSFWorkbook)workbook, dataValues.getFormulaOfLink());
	}

	private String getTitleFormula(LinkedDataRecord dataCategoryLabels) {
		if (dataCategoryLabels == null) {
			return "";
		}
		return HSSFFormulaParser.toFormulaString((HSSFWorkbook)workbook, dataCategoryLabels.getFormulaOfLink());
	}
	
	private String getChartTitle(HSSFSheet sheet, ChartInfo chartInfo) {
//		final boolean autoTitleDeleted = ((HSSFChart)chartInfo).isAutoTitleDeleted();
//		if (!autoTitleDeleted) {
//			final String title = ((HSSFChart)chartInfo).getChartTitle();
//			return title != null ? title :  "pie".equals(type) || "ring".equals(type) ? getFirstSeriesTitle(drawer, sheet, chartInfo) : "";
//		}
		return "";
	}

	@Override
	protected int getXoffsetInPixel(ClientAnchor anchor, Sheet poiSheet) {
	    final int firstColumn = anchor.getCol1();
	    final int firstXoffset = anchor.getDx1();
	    
	    final int columnWidthPixel = XUtils.getWidthAny(poiSheet,firstColumn, CHRACTER_WIDTH);
	    return firstXoffset >= 1024 ? columnWidthPixel : (int) Math.round(((double)columnWidthPixel) * firstXoffset / 1024);  
	}
	
	@Override
	protected int getYoffsetInPixel(ClientAnchor anchor, Sheet poiSheet) {
		final int firstRow = anchor.getRow1();
		final int firstYoffset = anchor.getDy1();

		final int rowHeightPixel = XUtils.getHeightAny(poiSheet,firstRow);
		return firstYoffset >= 256 ? rowHeightPixel : (int) Math.round(((double)rowHeightPixel) * firstYoffset / 256);  
	}
}
