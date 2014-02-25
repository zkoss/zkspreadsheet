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
package org.zkoss.zss.range.impl.imexp;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.usermodel.charts.*;
import org.zkoss.poi.ss.util.*;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.poi.xssf.usermodel.XSSFAutoFilter.XSSFFilterColumn;
import org.zkoss.poi.xssf.usermodel.charts.*;
import org.zkoss.zss.model.*;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SChart.BarDirection;
import org.zkoss.zss.model.SChart.ChartLegendPosition;
import org.zkoss.zss.model.SDataValidation.ErrorStyle;
import org.zkoss.zss.model.SDataValidation.OperatorType;
import org.zkoss.zss.model.SDataValidation.ValidationType;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.model.chart.*;
/**
 * 
 * @author dennis, kuro, Hawk
 * @since 3.5.0
 */
public class ExcelXlsxExporter extends AbstractExcelExporter {
	
	protected void exportColumnArray(SSheet sheet, Sheet poiSheet, SColumnArray columnArr) {
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
    	col.setWidth(UnitUtil.pxToCTChar(columnArr.getWidth(), AbstractExcelImporter.CHRACTER_WIDTH));
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
	protected void exportChart(SSheet sheet, Sheet poiSheet) {
		for (SChart chart: sheet.getCharts()){
			ChartData chartData = fillPoiChartData(chart);
			if (chartData != null){ //an unsupported chart has null chart data
				plotPoiChart(chart, chartData, sheet, poiSheet );
			}
		}
	}

	/**
	 * Reference DrawingManagerImpl.addPicture()
	 */
	@Override
	protected void exportPicture(SSheet sheet, Sheet poiSheet) {
		for (SPicture picture : sheet.getPictures()){
			int poiPictureIndex = workbook.addPicture(picture.getData(), toPoiPictureFormat(picture.getFormat()));
			poiSheet.createDrawingPatriarch().createPicture(toClientAnchor(picture.getAnchor(), sheet), poiPictureIndex);
		}
	}
	
	private int toPoiPictureFormat(Format format){
		switch(format){
		case GIF:
			return XSSFWorkbook.PICTURE_TYPE_GIF;
		case JPG:
			return Workbook.PICTURE_TYPE_JPEG;
		case PNG:
		default:
			return Workbook.PICTURE_TYPE_PNG;
		}
	}
	
	/**
	 * 
	 * @param chart
	 * @return a POI ChartData filled with Spreadsheet chart data, or null if the chart type is unsupported.   
	 */
	private ChartData fillPoiChartData(SChart chart) {
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
				fillXYZData((SGeneralChartData)chart.getData(), xyzData);
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
				fillXYData((SGeneralChartData)chart.getData(), xyData);
				chartData = xyData;
				break;
//			case STOCK: TODO XSSFStockChartData is implemented with errors.
//				categoryData = new XSSFStockChartData();
//				break;
			default:
				return chartData;
		}
		if (categoryData != null){
			fillCategoryData((SGeneralChartData)chart.getData(), categoryData);
			chartData = categoryData;
		}
		return chartData;
	}
	
	/**
	 * Create and plot a POI chart with its chart data.
	 * @param chart
	 * @param chartData
	 * @param sheet
	 * @param poiSheet the sheet where the POI chart locates
	 */
	private void plotPoiChart(SChart chart, ChartData chartData, SSheet sheet, Sheet poiSheet){
		Chart poiChart = poiSheet.createDrawingPatriarch().createChart(toClientAnchor(chart.getAnchor(),sheet));
		//TODO export a chart's title, no POI API supported
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
			poiChart.plot(chartData, bottomAxis, poiChart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT));
		} else {
			poiChart.plot(chartData);
		}
	}
	
	
	private ClientAnchor toClientAnchor(ViewAnchor viewAnchor, SSheet sheet){
		ViewAnchor rightBottomAnchor = viewAnchor.getRightBottomAnchor(sheet);
		
		ClientAnchor clientAnchor = new XSSFClientAnchor(UnitUtil.pxToEmu(viewAnchor.getXOffset()),UnitUtil.pxToEmu(viewAnchor.getYOffset()),
				UnitUtil.pxToEmu(rightBottomAnchor.getXOffset()),UnitUtil.pxToEmu(rightBottomAnchor.getYOffset()),
				viewAnchor.getColumnIndex(),viewAnchor.getRowIndex(),
				rightBottomAnchor.getColumnIndex(),rightBottomAnchor.getRowIndex());
		return clientAnchor;
	}
	
	/**
	 * reference ChartDataUtil.fillCategoryData()
	 * @param chart
	 * @param categoryData
	 */
	private void fillCategoryData(SGeneralChartData chartData, CategoryData categoryData){
		ChartDataSource<?> categories = createCategoryChartDataSource(chartData);
		for (int i=0 ; i < chartData.getNumOfSeries() ; i++){
			SSeries series = chartData.getSeries(i);
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
	private void fillXYData(SGeneralChartData chartData, XYData xyData){
		for (int i=0 ; i < chartData.getNumOfSeries() ; i++){
			final SSeries series = chartData.getSeries(i);
			ChartTextSource title = createChartTextSource(series);
			ChartDataSource<? extends Number> xValues = createXValueDataSource(series);
			ChartDataSource<? extends Number> yValues = createYValueDataSource(series);
			xyData.addSerie(title, xValues, yValues);
		}
	}

	/**
	 * reference ChartDataUtil.fillXYZData()
	 */
	private void fillXYZData(SGeneralChartData chartData, XYZData xyzData){
		for (int i=0 ; i < chartData.getNumOfSeries() ; i++){
			final SSeries series = chartData.getSeries(i);
			ChartTextSource title = createChartTextSource(series);
			ChartDataSource<? extends Number> xValues = createXValueDataSource(series);
			ChartDataSource<? extends Number> yValues = createYValueDataSource(series);
			ChartDataSource<? extends Number> zValues = createZValueDataSource(series);
			xyzData.addSerie(title, xValues, yValues, zValues);
		}
	}
	
	private ChartDataSource<Number> createXValueDataSource(final SSeries series) {
		return new ChartDataSource<Number>() {

			@Override
			public int getPointCount() {
				return series.getNumOfXValue();
			}

			@Override
			public Number getPointAt(int index) {
				try{
					return Double.parseDouble(series.getXValue(index).toString());
				}catch(NumberFormatException nfe){
					return index;
				}
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

	private ChartDataSource<Number> createYValueDataSource(final SSeries series) {
		return new ChartDataSource<Number>() {

			@Override
			public int getPointCount() {
				return series.getNumOfYValue();
			}

			@Override
			public Number getPointAt(int index) {
				try{
					return Double.parseDouble(series.getYValue(index).toString());
				}catch (NumberFormatException nfe) {
					return index;
				}
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
	
	private ChartDataSource<Number> createZValueDataSource(final SSeries series) {
		return new ChartDataSource<Number>() {

			@Override
			public int getPointCount() {
				return series.getNumOfZValue();
			}

			@Override
			public Number getPointAt(int index) {
				try{
					return Double.parseDouble(series.getZValue(index).toString());
				}catch (NumberFormatException e) {
					return index;
				}
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
	
	private ChartTextSource createChartTextSource(final SSeries series){
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

	private ChartDataSource<?> createCategoryChartDataSource(final SGeneralChartData chartData){
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
	
	private ChartGrouping toPoiGrouping(org.zkoss.zss.model.SChart.ChartGrouping grouping){
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
	
	private ChartDirection toPoiBarDirection(BarDirection direction){
		switch(direction){
			case VERTICAL:
				return ChartDirection.VERTICAL;
			case HORIZONTAL:
			default:
				return ChartDirection.HORIZONTAL;
		}
		
	}
	
	private LegendPosition toPoiLegendPosition(ChartLegendPosition position){
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

	/**
	 * According to {@link ValidationType}, FORMULA means custom validation.
	 */
	@Override
	protected void exportValidation(SSheet sheet, Sheet poiSheet) {
		for (SDataValidation validation : sheet.getDataValidations()){
			int operatorType = toPoiOperatorType(validation.getOperatorType());
			String formula1 = validation.getValue1Formula();
			String formula2 = validation.getValue2Formula();
			DataValidationConstraint constraint = null;
			switch(validation.getValidationType()){
				case TIME:
					constraint = poiSheet.getDataValidationHelper().createTimeConstraint(operatorType, formula1, formula2);
					break;
				case TEXT_LENGTH:
					constraint = poiSheet.getDataValidationHelper().createTextLengthConstraint(operatorType, formula1, formula2);
					break;
				case DATE:
					//the last argument, dateFormat, is only used in XLS. We just pass empty string here. 
					constraint = poiSheet.getDataValidationHelper().createDateConstraint(operatorType, formula1, formula2, "");
					break;
				case LIST:
					constraint = poiSheet.getDataValidationHelper().createFormulaListConstraint(formula1);
					break;
				case INTEGER:
					constraint = poiSheet.getDataValidationHelper().createIntegerConstraint(operatorType, formula1, formula2);
					break;
				case FORMULA: // custom
					constraint = poiSheet.getDataValidationHelper().createCustomConstraint(formula1);
					break;
				case DECIMAL:
					constraint = poiSheet.getDataValidationHelper().createDecimalConstraint(operatorType, formula1, formula2);
					break;
				case ANY:
					// ANY validation type means no validation
				default:
					continue;
			}
			CellRegion firstRegion = validation.getRegion();
			
			DataValidation poiValidation = poiSheet.getDataValidationHelper().createValidation(constraint, 
					new CellRangeAddressList(firstRegion.getRow(),firstRegion.getLastRow(), firstRegion.getColumn(), firstRegion.getLastColumn()));
			
			poiValidation.setEmptyCellAllowed(validation.isEmptyCellAllowed());
			poiValidation.setSuppressDropDownArrow(validation.isShowDropDownArrow());
			
			poiValidation.setErrorStyle(toPoiErrorStyle(validation.getErrorStyle()));
			poiValidation.createErrorBox(validation.getErrorBoxTitle(), validation.getErrorBoxText());
			poiValidation.setShowErrorBox(validation.isShowErrorBox());
			
			poiValidation.createPromptBox(validation.getPromptBoxTitle(), validation.getPromptBoxText());
			poiValidation.setShowPromptBox(validation.isShowPromptBox());
			
			poiSheet.addValidationData(poiValidation);
		}
	}
	
	private int toPoiOperatorType(OperatorType type){
		switch (type) {
			case NOT_EQUAL:
				return DataValidationConstraint.OperatorType.NOT_EQUAL;
			case NOT_BETWEEN:	
				return DataValidationConstraint.OperatorType.NOT_BETWEEN;
			case LESS_THAN:
				return DataValidationConstraint.OperatorType.LESS_THAN;
			case LESS_OR_EQUAL:
				return DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
			case GREATER_THAN:
				return DataValidationConstraint.OperatorType.GREATER_THAN;
			case GREATER_OR_EQUAL:
				return DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
			case EQUAL:
				return DataValidationConstraint.OperatorType.EQUAL;
			case BETWEEN:
			default:
				return DataValidationConstraint.OperatorType.BETWEEN;
		}
	}
	
	private int toPoiErrorStyle(ErrorStyle style){
		switch(style){
		case INFO:
			return DataValidation.ErrorStyle.INFO;
		case WARNING:
			return DataValidation.ErrorStyle.WARNING;
		case STOP:
		default:
			return DataValidation.ErrorStyle.STOP;
		}
	}

	/**
	 * See Javadoc at {@link AbstractExcelImporter} importAutoFilter().
	 */
	@SuppressWarnings("unchecked")
	@Override
	protected void exportAutoFilter(SSheet sheet, Sheet poiSheet) {
		SAutoFilter autoFilter = sheet.getAutoFilter();
		if (autoFilter != null){
			CellRegion region = autoFilter.getRegion();
			XSSFAutoFilter poiAutoFilter = (XSSFAutoFilter)poiSheet.setAutoFilter(new CellRangeAddress(region.getRow(), region.getLastRow(), region.getColumn(), region.getLastColumn()));
			int numberOfColumn = region.getLastColumn() - region.getColumn() + 1;
			for( int i = 0 ; i < numberOfColumn ; i++){
				NFilterColumn srcFilterColumn = autoFilter.getFilterColumn(i, false);
				if (srcFilterColumn == null){
					continue;
				}
				XSSFFilterColumn destFilterColumn = (XSSFFilterColumn)poiAutoFilter.getOrCreateFilterColumn(i);
				Object[] criteria1 = null;
				if (srcFilterColumn.getCriteria1()!=null){
					criteria1 = srcFilterColumn.getCriteria1().toArray(new String[0]);
				}
				Object[] criteria2 = null;
				if (srcFilterColumn.getCriteria1()!=null){
					criteria2 = srcFilterColumn.getCriteria2().toArray(new String[0]);
				}
				destFilterColumn.setProperties(criteria1, ExporterEnumUtil.toPoiFilterOperator(srcFilterColumn.getOperator()),
						criteria2, srcFilterColumn.isShowButton());
			}
		}
	}
	
}
