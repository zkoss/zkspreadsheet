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

import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.usermodel.charts.*;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.poi.xssf.usermodel.charts.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NChart.NBarDirection;
import org.zkoss.zss.ngmodel.NChart.NChartGrouping;
import org.zkoss.zss.ngmodel.NChart.NChartLegendPosition;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.chart.*;
/**
 * 
 * @author dennis, kuro
 * @since 3.5.0
 */
public class NExcelXlsxExporter extends AbstractExcelExporter {
	
	protected void exportColumnArray(NSheet sheet, Sheet poiSheet, NColumnArray columnArr) {
		XSSFSheet xssfSheet = (XSSFSheet) poiSheet;
		
        CTWorksheet ctSheet = xssfSheet.getCTWorksheet();
    	if(xssfSheet.getCTWorksheet().sizeOfColsArray() == 0) {
    		xssfSheet.getCTWorksheet().addNewCols();
    	}
    		
    	CTCol col = ctSheet.getColsArray(0).addNewCol();
        col.setMin(columnArr.getIndex()+1);
        col.setMax(columnArr.getLastIndex()+1);
    	col.setStyle(toPOICellStyle(columnArr.getCellStyle()).getIndex());
    	col.setCustomWidth(true);
    	col.setWidth(XUtils.pxToCTChar(columnArr.getWidth(), AbstractExcelImporter.CHRACTER_WIDTH));
    	col.setHidden(columnArr.isHidden());
	}

	@Override
	protected Workbook createPoiBook() {
		return new XSSFWorkbook();
	}

	/**
	 * reference DrawingManagerImpl.addChartX()
	 */
	@Override
	protected void exportChart(NSheet sheet, Sheet poiSheet) {
		for (NChart chart: sheet.getCharts()){
			ChartData chartData = fillChartData(chart);
			plotChart(chart, chartData, sheet, poiSheet );
		}
	}

	private ChartData fillChartData(NChart chart) {
		CategoryData categoryData = null;
		ChartData chartData = null;
		switch(chart.getType()){
			case AREA:
				if (chart.isThreeD()){
					categoryData = new XSSFArea3DChartData();
					((XSSFArea3DChartData)categoryData).setGrouping(toPoiGrouping(chart.getGrouping()));
				}else{
					categoryData = new XSSFAreaChartData();
					((XSSFAreaChartData)categoryData).setGrouping(toPoiGrouping(chart.getGrouping()));
				}
				break;
			case BAR:
				if (chart.isThreeD()){
					categoryData = new XSSFBar3DChartData();				
					((XSSFBar3DChartData)categoryData).setGrouping(toPoiGrouping(chart.getGrouping()));
					((XSSFBar3DChartData)categoryData).setBarDirection(toPoiBarDirection(chart.getBarDirection()));
				}else{
					categoryData = new XSSFBarChartData();
					((XSSFBarChartData)categoryData).setGrouping(toPoiGrouping(chart.getGrouping()));
					((XSSFBarChartData)categoryData).setBarDirection(toPoiBarDirection(chart.getBarDirection()));
				}
				break;
			case BUBBLE:
				XYZData xyzData  = new XSSFBubbleChartData();
				fillXYZData(chart, xyzData);
				chartData = xyzData;
				break;
			case COLUMN:
				if (chart.isThreeD()){
					categoryData = new XSSFColumn3DChartData();
					((XSSFColumn3DChartData)categoryData).setGrouping(toPoiGrouping(chart.getGrouping()));
					((XSSFColumn3DChartData)categoryData).setBarDirection(toPoiBarDirection(chart.getBarDirection()));
				}else{
					categoryData = new XSSFColumnChartData();
					((XSSFColumnChartData)categoryData).setGrouping(toPoiGrouping(chart.getGrouping()));
					((XSSFColumnChartData)categoryData).setBarDirection(toPoiBarDirection(chart.getBarDirection()));
				}
				break;
			case DOUGHNUT:
				categoryData = new XSSFDoughnutChartData();
				break;
			case LINE:
				if (chart.isThreeD()){
					categoryData = new XSSFLine3DChartData();
				}else{
					categoryData = new XSSFLineChartData();
				}
				break;
			case PIE:
				if (chart.isThreeD()){
					categoryData = new XSSFPie3DChartData();
				}else{
					categoryData = new XSSFPieChartData();
				}
				break;
			case SCATTER:
				XYData xyData =  new XSSFScatChartData();
				fillXYData(chart, xyData);
				chartData = xyData;
				break;
//				case STOCK: TODO XSSFStockChartData is implemented with errors.
//					categoryData = new XSSFStockChartData();
//					break;
			default:
			return chartData;
		}
		if (categoryData != null){
			fillCategoryData(chart, categoryData);
			chartData = categoryData;
		}
		return chartData;
	}
	
	private void plotChart(NChart chart, ChartData chartData, NSheet sheet, Sheet poiSheet){
		final Drawing drawing = poiSheet.createDrawingPatriarch();
		ClientAnchor anchor = toClientAnchor(chart.getAnchor(),sheet, poiSheet);
		final Chart poiChart = drawing.createChart(anchor);
		//TODO export a chart's title
		if (chart.isThreeD()){
			poiChart.getOrCreateView3D();
		}
		if (chart.getLegendPosition() != null) {
			ChartLegend legend = poiChart.getOrCreateLegend();
			legend.setPosition(toPoiLegendPosition(chart.getLegendPosition()));
		}
		ChartAxis bottomAxis = null;
		switch(chart.getType()) {
			case AREA:
			case BAR:
			case COLUMN:
			case LINE:
				bottomAxis=  poiChart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
				break;
			case BUBBLE:
			case SCATTER:
				bottomAxis = poiChart.getChartAxisFactory().createValueAxis(AxisPosition.BOTTOM);
				break;
		}
		if (bottomAxis != null) {
			bottomAxis.setCrosses(AxisCrosses.AUTO_ZERO);
			final ValueAxis leftAxis = poiChart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);
			leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
			poiChart.plot(chartData, bottomAxis, leftAxis);
		} else {
			poiChart.plot(chartData);
		}
	}
	
	private ClientAnchor toClientAnchor(NViewAnchor viewAnchor, NSheet sheet, Sheet poiSheet){
		
		int dx1 = viewAnchor.getXOffset();
		int offsetPlusChartWidth = dx1 + viewAnchor.getWidth();
		int lastColumn = viewAnchor.getColumnIndex();
		int dx2 = 0;
		//minus width of each inter-column to find last column index and dx2
		for (int column = viewAnchor.getColumnIndex(); ;column++){
			int interColumnWidth = sheet.getColumn(column).getWidth();
			if (offsetPlusChartWidth - interColumnWidth < 0){ 
				lastColumn = column;
				dx2 = offsetPlusChartWidth;
				break;
			}else{
				offsetPlusChartWidth -= interColumnWidth;
			}
		}
		
		int dy1 = viewAnchor.getYOffset();
		int offsetPlusChartHeight = dy1 + viewAnchor.getHeight();
		int lastRow = viewAnchor.getRowIndex();
		int dy2 = 0;
		//minus height of each inter-row to find last row index and dy2
		for (int row = viewAnchor.getRowIndex(); ;row++){
			int interRowHeight = sheet.getRow(row).getHeight();
			if (offsetPlusChartHeight - interRowHeight < 0){
				lastRow = row;
				dy2 = offsetPlusChartHeight;
				break;
			}else{
				offsetPlusChartHeight -= interRowHeight;
			}
		}
		
		ClientAnchor clientAnchor = new XSSFClientAnchor(UnitUtil.pxToEmu(dx1),UnitUtil.pxToEmu(dy1),
				UnitUtil.pxToEmu(dx2),UnitUtil.pxToEmu(dy2),
				viewAnchor.getColumnIndex(),viewAnchor.getRowIndex(),
				lastColumn,lastRow);
		return clientAnchor;
	}
	
	/**
	 * reference ChartDataUtil.fillCategoryData()
	 * @param chart
	 * @param categoryData
	 */
	private void fillCategoryData(NChart chart, CategoryData categoryData){
		NGeneralChartData chartData = (NGeneralChartData)chart.getData();
		ChartDataSource<?> categories = createCategoryChartDataSource(chartData);
		for (int i=0 ; i < chartData.getNumOfSeries() ; i++){
			NSeries series = chartData.getSeries(i);
			ChartTextSource title = createChartTextSource(series);
			ChartDataSource<? extends Number> values = createXValueDataSource(series);
			categoryData.addSerie(title, categories, values);
		}
	}
	
	/**
	 * reference ChartDataUtil.fillXYData()
	 * @param chart
	 * @param xyData
	 */
	private void fillXYData(NChart chart, XYData xyData){
		NGeneralChartData chartData = (NGeneralChartData)chart.getData();
		for (int i=0 ; i < chartData.getNumOfSeries() ; i++){
			final NSeries series = chartData.getSeries(i);
			ChartTextSource title = createChartTextSource(series);
			ChartDataSource<? extends Number> xValues = createXValueDataSource(series);
			ChartDataSource<? extends Number> yValues = createYValueDataSource(series);
			xyData.addSerie(title, xValues, yValues);
		}
	}

	/**
	 * reference ChartDataUtil.fillXYZData()
	 */
	private void fillXYZData(NChart chart, XYZData xyzData){
		NGeneralChartData chartData = (NGeneralChartData)chart.getData();
		for (int i=0 ; i < chartData.getNumOfSeries() ; i++){
			final NSeries series = chartData.getSeries(i);
			ChartTextSource title = createChartTextSource(series);
			ChartDataSource<? extends Number> xValues = createXValueDataSource(series);
			ChartDataSource<? extends Number> yValues = createYValueDataSource(series);
			ChartDataSource<? extends Number> zValues = createZValueDataSource(series);
			xyzData.addSerie(title, xValues, yValues, zValues);
		}
	}
	
	private ChartDataSource<Number> createXValueDataSource(final NSeries series) {
		return new ChartDataSource<Number>() {

			@Override
			public int getPointCount() {
				return series.getNumOfXValue();
			}

			@Override
			public Number getPointAt(int index) {
				return Double.parseDouble(series.getXValue(index).toString());
			}

			@Override
			public boolean isReference() {
				return true;
			}

			@Override
			public boolean isNumeric() {
				return true;
			}

			@Override
			public String getFormulaString() {
				return series.getXValuesFormula();
			}

			@Override
			public void renameSheet(String oldname, String newname) {
			}
		};
	}

	private ChartDataSource<Number> createYValueDataSource(final NSeries series) {
		return new ChartDataSource<Number>() {

			@Override
			public int getPointCount() {
				return series.getNumOfYValue();
			}

			@Override
			public Number getPointAt(int index) {
				return Double.parseDouble(series.getYValue(index).toString());
			}

			@Override
			public boolean isReference() {
				return true;
			}

			@Override
			public boolean isNumeric() {
				return true;
			}

			@Override
			public String getFormulaString() {
				return series.getYValuesFormula();
			}

			@Override
			public void renameSheet(String oldname, String newname) {
			}
		};
	}
	
	private ChartDataSource<Number> createZValueDataSource(final NSeries series) {
		return new ChartDataSource<Number>() {

			@Override
			public int getPointCount() {
				return series.getNumOfZValue();
			}

			@Override
			public Number getPointAt(int index) {
				return Double.parseDouble(series.getZValue(index).toString());
			}

			@Override
			public boolean isReference() {
				return true;
			}

			@Override
			public boolean isNumeric() {
				return true;
			}

			@Override
			public String getFormulaString() {
				return series.getZValuesFormula();
			}

			@Override
			public void renameSheet(String oldname, String newname) {
			}
		};
	}
	
	private ChartTextSource createChartTextSource(final NSeries series){
		return new ChartTextSource() {
			
			@Override
			public void renameSheet(String oldname, String newname) {
			}
			
			@Override
			public boolean isReference() {
				return true;
			}
			
			@Override
			public String getTextString() {
				return series.getName();
			}
			
			@Override
			public String getFormulaString() {
				return series.getNameFormula();
			}
		};
		
	}

	private ChartDataSource<?> createCategoryChartDataSource(final NGeneralChartData chartData){
		return new ChartDataSource<String>() {

			@Override
			public int getPointCount() {
				return chartData.getNumOfCategory();
			}

			@Override
			public String getPointAt(int index) {
				return chartData.getCategory(index).toString();
			}

			@Override
			public boolean isReference() {
				return true;
			}

			@Override
			public boolean isNumeric() {
				return false;
			}

			@Override
			public String getFormulaString() {
				return chartData.getCategoriesFormula();
			}

			@Override
			public void renameSheet(String oldname, String newname) {
			}
		};
	}
	
	private ChartGrouping toPoiGrouping(NChartGrouping grouping){
		switch(grouping){
			case CLUSTERED:
				return ChartGrouping.CLUSTERED;
			case PERCENT_STACKED:
				return ChartGrouping.PERCENT_STACKED;
			case STACKED:
				return ChartGrouping.STACKED;
			case STANDARD:
			default:
				return ChartGrouping.STANDARD;
		}
	}
	
	private ChartDirection toPoiBarDirection(NBarDirection direction){
		switch(direction){
			case VERTICAL:
				return ChartDirection.VERTICAL;
			case HORIZONTAL:
			default:
				return ChartDirection.HORIZONTAL;
		}
		
	}
	
	private LegendPosition toPoiLegendPosition(NChartLegendPosition position){
		switch(position){
			case BOTTOM:
				return LegendPosition.BOTTOM;
			case TOP:
				return LegendPosition.TOP;
			case TOP_RIGHT:
				return LegendPosition.TOP_RIGHT;
			case LEFT:
				return LegendPosition.LEFT;
			case RIGHT:
			default:
				return LegendPosition.RIGHT;
			
		}
	}
}
