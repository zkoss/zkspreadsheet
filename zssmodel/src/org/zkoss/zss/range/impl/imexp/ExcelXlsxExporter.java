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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.usermodel.charts.*;
import org.zkoss.poi.ss.util.*;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.poi.xssf.usermodel.XSSFAutoFilter.XSSFFilterColumn;
import org.zkoss.poi.xssf.usermodel.XSSFTableColumn.TotalsRowFunction;
import org.zkoss.poi.xssf.usermodel.charts.*;
import org.zkoss.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.zkoss.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.zkoss.poi.xssf.usermodel.XSSFRichTextString;
import org.zkoss.zss.model.*;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SBorder.BorderType;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.SRichText.Segment;
import org.zkoss.zss.model.STableColumn.STotalsRowFunction;
import org.zkoss.zss.model.SDataValidation.ValidationType;
import org.zkoss.zss.model.chart.*;
import org.zkoss.zss.model.impl.AbstractBookAdv;
import org.zkoss.zss.model.impl.AbstractDataValidationAdv;
import org.zkoss.zss.model.impl.AbstractFillAdv;
import org.zkoss.zss.model.impl.AbstractFontAdv;
import org.zkoss.zss.model.impl.AbstractColorAdv;
import org.zkoss.zss.model.impl.SheetImpl;
/**
 * 
 * @author dennis, kuro, Hawk
 * @since 3.5.0
 */
public class ExcelXlsxExporter extends AbstractExcelExporter {
	private static final long serialVersionUID = 20141231175402L;

	protected void exportColumnArray(SSheet sheet, Sheet poiSheet, SColumnArray columnArr) {
		XSSFSheet xssfSheet = (XSSFSheet) poiSheet;
		
        CTWorksheet ctSheet = xssfSheet.getCTWorksheet();
    	if(xssfSheet.getCTWorksheet().sizeOfColsArray() == 0) {
    		xssfSheet.getCTWorksheet().addNewCols();
    	}

    	//ZSS-1132
		final int defaultWidth = sheet.getDefaultColumnWidth();
		final int columnWidth = columnArr.getWidth();
		final AbstractBookAdv book = (AbstractBookAdv)sheet.getBook();
		final int charWidth = book.getCharWidth();
		
    	CTCol col = ctSheet.getColsArray(0).addNewCol();
        col.setMin(columnArr.getIndex()+1);
        col.setMax(columnArr.getLastIndex()+1);
    	col.setStyle(toPOICellStyle(columnArr.getCellStyle()).getIndex());
    	col.setCustomWidth(columnArr.isCustomWidth() || columnWidth != defaultWidth); //ZSS-1132
    	col.setWidth(UnitUtil.pxToCTChar(columnWidth, charWidth));
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
			int poiPictureIndex = exportedPicDataMap.get(picture.getPictureData().getIndex()); //ZSS-735
			poiSheet.createDrawingPatriarch().createPicture(toClientAnchor(picture.getAnchor(), sheet), poiPictureIndex);
		}
	}
	
	/**
	 * 
	 * @param chart
	 * @return a POI ChartData filled with Spreadsheet chart data, or null if the chart type is unsupported.   
	 */
	protected ChartData fillPoiChartData(SChart chart) {
		CategoryData categoryData = null;
		ChartData chartData = null;
		switch(chart.getType()){
			case AREA:
				if (chart.isThreeD()){
					categoryData = new XSSFArea3DChartData();
					((XSSFArea3DChartData)categoryData).setGrouping(PoiEnumConversion.toPoiGrouping(chart.getGrouping()));
				}else{
					categoryData = new XSSFAreaChartData();
					((XSSFAreaChartData)categoryData).setGrouping(PoiEnumConversion.toPoiGrouping(chart.getGrouping()));
				}
				break;
			case BAR:
				if (chart.isThreeD()){
					categoryData = new XSSFBar3DChartData();				
					((XSSFBar3DChartData)categoryData).setGrouping(PoiEnumConversion.toPoiGrouping(chart.getGrouping()));
					((XSSFBar3DChartData)categoryData).setBarDirection(PoiEnumConversion.toPoiBarDirection(chart.getBarDirection()));
				}else{
					categoryData = new XSSFBarChartData();
					((XSSFBarChartData)categoryData).setGrouping(PoiEnumConversion.toPoiGrouping(chart.getGrouping()));
					((XSSFBarChartData)categoryData).setBarDirection(PoiEnumConversion.toPoiBarDirection(chart.getBarDirection()));
					((XSSFBarChartData)categoryData).setBarOverlap(chart.getBarOverlap()); //ZSS-830
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
					((XSSFColumn3DChartData)categoryData).setGrouping(PoiEnumConversion.toPoiGrouping(chart.getGrouping()));
					((XSSFColumn3DChartData)categoryData).setBarDirection(PoiEnumConversion.toPoiBarDirection(chart.getBarDirection()));
				}else{
					categoryData = new XSSFColumnChartData();
					((XSSFColumnChartData)categoryData).setGrouping(PoiEnumConversion.toPoiGrouping(chart.getGrouping()));
					((XSSFColumnChartData)categoryData).setBarDirection(PoiEnumConversion.toPoiBarDirection(chart.getBarDirection()));
					((XSSFColumnChartData)categoryData).setBarOverlap(chart.getBarOverlap());
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
	protected void plotPoiChart(SChart chart, ChartData chartData, SSheet sheet, Sheet poiSheet){
		Chart poiChart = poiSheet.createDrawingPatriarch().createChart(toClientAnchor(chart.getAnchor(),sheet));
		//TODO export a chart's title, no POI API supported
		if (chart.isThreeD()){
			//ZSS-830
			XSSFView3D view3d = (XSSFView3D) poiChart.getOrCreateView3D();
			if (chart.getRotX() != 0) view3d.setRotX(chart.getRotX());
			if (chart.getRotY() != 0) view3d.setRotY(chart.getRotY());
			if (chart.getPerspective() != 30) view3d.setPerspective(chart.getPerspective());
			if (chart.getHPercent() != 100) view3d.setHPercent(chart.getHPercent());
			if (chart.getDepthPercent() != 100) view3d.setDepthPercent(chart.getDepthPercent());
			if (!chart.isRightAngleAxes()) view3d.setRightAngleAxes(false);
		}
		if (chart.getLegendPosition() != null) {
			ChartLegend legend = poiChart.getOrCreateLegend();
			legend.setPosition(PoiEnumConversion.toPoiLegendPosition(chart.getLegendPosition()));
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
		
		poiChart.setPlotOnlyVisibleCells(chart.isPlotOnlyVisibleCells());
	}
	
	
	protected ClientAnchor toClientAnchor(ViewAnchor viewAnchor, SSheet sheet){
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
	protected void fillCategoryData(SGeneralChartData chartData, CategoryData categoryData){
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
	protected void fillXYData(SGeneralChartData chartData, XYData xyData){
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
	protected void fillXYZData(SGeneralChartData chartData, XYZData xyzData){
		for (int i=0 ; i < chartData.getNumOfSeries() ; i++){
			final SSeries series = chartData.getSeries(i);
			ChartTextSource title = createChartTextSource(series);
			ChartDataSource<? extends Number> xValues = createXValueDataSource(series);
			ChartDataSource<? extends Number> yValues = createYValueDataSource(series);
			ChartDataSource<? extends Number> zValues = createZValueDataSource(series);
			xyzData.addSerie(title, xValues, yValues, zValues);
		}
	}
	
	protected ChartDataSource<Number> createXValueDataSource(final SSeries series) {
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

	protected ChartDataSource<Number> createYValueDataSource(final SSeries series) {
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
	
	protected ChartDataSource<Number> createZValueDataSource(final SSeries series) {
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
	
	protected ChartTextSource createChartTextSource(final SSeries series){
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

	protected ChartDataSource<?> createCategoryChartDataSource(final SGeneralChartData chartData){
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
	
	/**
	 * According to {@link ValidationType}, FORMULA means custom validation.
	 */
	@Override
	protected void exportValidation(SSheet sheet, Sheet poiSheet) {
		for (SDataValidation validation : sheet.getDataValidations()){
			int operatorType = PoiEnumConversion.toPoiOperatorType(validation.getOperatorType());
			String formula1 = ((AbstractDataValidationAdv)validation).getEscapedFormula1();
			String formula2 = ((AbstractDataValidationAdv)validation).getEscapedFormula2();
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
				case CUSTOM: // custom
					constraint = poiSheet.getDataValidationHelper().createCustomConstraint(formula1);
					break;
				case DECIMAL:
					constraint = poiSheet.getDataValidationHelper().createDecimalConstraint(operatorType, formula1, formula2);
					break;
				case ANY:
					constraint = poiSheet.getDataValidationHelper().createAnyConstraint(); //ZSS-835
					break;
				default:
					continue;
			}
			if (!validation.getRegions().isEmpty()) { // ZSS-835
				final CellRangeAddressList rgnList = new CellRangeAddressList();
				for (CellRegion rgn : validation.getRegions()) { // must prepare rgnList then create poiValidation
					rgnList.addCellRangeAddress(rgn.getRow(), rgn.getColumn(), rgn.getLastRow(), rgn.getLastColumn());
				}
				DataValidation poiValidation = 
					poiSheet.getDataValidationHelper().createValidation(constraint, rgnList);
				
				poiValidation.setEmptyCellAllowed(validation.isIgnoreBlank());
				poiValidation.setSuppressDropDownArrow(validation.isInCellDropdown());
				
				poiValidation.setErrorStyle(PoiEnumConversion.toPoiErrorStyle(validation.getAlertStyle()));
				poiValidation.createErrorBox(validation.getErrorTitle(), validation.getErrorMessage());
				poiValidation.setShowErrorBox(validation.isShowError());
				
				poiValidation.createPromptBox(validation.getInputTitle(), validation.getInputMessage());
				poiValidation.setShowPromptBox(validation.isShowInput());
				
				poiSheet.addValidationData(poiValidation);
			}
		}
	}
	
	/**
	 * See Javadoc at {@link AbstractExcelImporter} importAutoFilter().
	 */
	@Override
	protected void exportAutoFilter(SSheet sheet, Sheet poiSheet) {
		SAutoFilter autoFilter = sheet.getAutoFilter();
		if (autoFilter != null){
			CellRegion region = autoFilter.getRegion();
			XSSFAutoFilter poiAutoFilter = (XSSFAutoFilter)poiSheet.setAutoFilter(new CellRangeAddress(region.getRow(), region.getLastRow(), region.getColumn(), region.getLastColumn()));
			int numberOfColumn = region.getLastColumn() - region.getColumn() + 1;
			exportFilterColumns(poiAutoFilter, autoFilter, numberOfColumn);
		}
	}
	
	//ZSS-1019
	protected void exportFilterColumns(XSSFAutoFilter poiAutoFilter, SAutoFilter autoFilter, int numberOfColumn) {
		final Map<String, Object> extra = new HashMap<String, Object>();
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
			
			//ZSS-1191
			final SColorFilter colorFilter = srcFilterColumn.getColorFilter();
			XSSFColorFilter poiFilter = null;
			if (colorFilter != null) {
				final SExtraStyle extraStyle = colorFilter.getExtraStyle();
				addPOIDxfCellStyle(extraStyle);
				final XSSFDxfCellStyle poiCellStyle = 
						(XSSFDxfCellStyle) styleTable.get(extraStyle);
				poiFilter = new XSSFColorFilter(poiCellStyle, colorFilter.isByFontColor());
			}
			extra.put("colorFilter", poiFilter);
			
			//ZSS-1224
			final SCustomFilters customFilters = srcFilterColumn.getCustomFilters();
			if (customFilters != null) {
				XSSFCustomFilters poiCustomFilters = new XSSFCustomFilters(destFilterColumn);
				poiCustomFilters.setAnd(customFilters.isAnd());
				final SCustomFilter srcFilter1 = customFilters.getCustomFilter1();
				final SCustomFilter srcFilter2 = customFilters.getCustomFilter2();
				poiCustomFilters.addCustomFilter(toPOIOpertor(srcFilter1.getOperator()), srcFilter1.getValue());
				if (srcFilter2 != null) {
					poiCustomFilters.addCustomFilter(toPOIOpertor(srcFilter2.getOperator()), srcFilter2.getValue());
				}
				extra.put("customFilters", poiCustomFilters); //ZSS-1224
			}
			
			//ZSS-1226
			final SDynamicFilter dynamicFilter = srcFilterColumn.getDynamicFilter();
			if (dynamicFilter != null) {
				XSSFDynamicFilter poiDynamicFilter = new XSSFDynamicFilter(destFilterColumn);
				poiDynamicFilter.setProperties(dynamicFilter.getMaxValue(), dynamicFilter.getValue(), dynamicFilter.getType()); //ZSS-1234
				extra.put("dynamicFilter", poiDynamicFilter);
			}
			
			//ZSS-1227
			final STop10Filter top10Filter = srcFilterColumn.getTop10Filter();
			if (top10Filter != null) {
				XSSFTop10Filter poiTop10Filter = new XSSFTop10Filter(destFilterColumn);
				poiTop10Filter.setProperties(top10Filter.isTop(), top10Filter.getValue(), top10Filter.isPercent(), top10Filter.getFilterValue());
				extra.put("top10Filter", poiTop10Filter);
			}
			
			//ZSS-1191
			destFilterColumn.setProperties(criteria1, PoiEnumConversion.toPoiFilterOperator(srcFilterColumn.getOperator()),
					criteria2, srcFilterColumn.isShowButton(), extra);
			
		}
	}
	
	//ZSS-1224
	private CustomFilter.Operator toPOIOpertor(FilterOp op) {
		return CustomFilter.Operator.valueOf(op.name());
	}
	
	/**
	 * Export hashed password directly to poiSheet.
	 */
	@Override
	protected void exportPassword(SSheet sheet, Sheet poiSheet) {
		
		//ZSS-1063
		final String hashValue = ((SheetImpl)sheet).getHashValue();
		if (hashValue != null) {
			final String saltValue = ((SheetImpl)sheet).getSaltValue();
			final String spinCount = ((SheetImpl)sheet).getSpinCount();
			final String algName = ((SheetImpl)sheet).getAlgName();
			
			((XSSFSheet)poiSheet).setHashValue(hashValue);
			((XSSFSheet)poiSheet).setSaltValue(saltValue);
			((XSSFSheet)poiSheet).setSpinCount(spinCount);
			((XSSFSheet)poiSheet).setAlgName(algName);
		} else {
			short hashpass = sheet.getHashedPassword();
			if (hashpass != 0) {
				((XSSFSheet)poiSheet).setPasswordHash(hashpass);
			}
		}
	}
	
	
	//ZSS-854 
	@Override
	protected CellStyle toPOIDefaultCellStyle(SCellStyle cellStyle) {
		//set Border
		short bottom = PoiEnumConversion.toPoiBorderType(cellStyle.getBorderBottom());
		short left = PoiEnumConversion.toPoiBorderType(cellStyle.getBorderLeft());
		short right = PoiEnumConversion.toPoiBorderType(cellStyle.getBorderRight());
		short top = PoiEnumConversion.toPoiBorderType(cellStyle.getBorderTop());
		Color bottomColor = toPOIColor(cellStyle.getBorderBottomColor());
		Color leftColor = toPOIColor(cellStyle.getBorderLeftColor());
		Color rightColor = toPOIColor(cellStyle.getBorderRightColor());
		Color topColor = toPOIColor(cellStyle.getBorderTopColor());
		CTBorder ct = CTBorder.Factory.newInstance();
		XSSFCellBorder border = new XSSFCellBorder(ct);
		border.prepareBorder(left, leftColor, top, topColor, right, rightColor, bottom, bottomColor);
		
		// fill
		//ZSS-857: SOLID pattern; switch fgColor and bgColor 
		SColor fgColor = cellStyle.getFillColor();
		SColor bgColor = cellStyle.getBackColor();
		if (cellStyle.getFillPattern() == FillPattern.SOLID) {
			SColor tmp = fgColor;
			fgColor = bgColor;
			bgColor = tmp;
		}
		Color fillColor = toPOIColor(fgColor);
		Color backColor = toPOIColor(bgColor);
		short pattern = PoiEnumConversion.toPoiFillPattern(cellStyle.getFillPattern());
		CTFill ctf = CTFill.Factory.newInstance();
		XSSFCellFill fill = new XSSFCellFill(ctf);
		fill.prepareFill(fillColor, backColor, pattern);
		
		// font
		XSSFFont font = (XSSFFont)toPOIFont(cellStyle.getFont());
		
		// refer from BookHelper#setDataFormat
		DataFormat df = workbook.createDataFormat();
		short fmt = df.getFormat(cellStyle.getDataFormat());
		
		XSSFCellStyle poiCellStyle = (XSSFCellStyle) ((XSSFWorkbook)workbook).createDefaultCellStyle(border, fill, font, fmt);

		//cell Alignment
		short hAlign = PoiEnumConversion.toPoiHorizontalAlignment(cellStyle.getAlignment());
		short vAlign = PoiEnumConversion.toPoiVerticalAlignment(cellStyle.getVerticalAlignment());
		boolean wrapText = cellStyle.isWrapText();
		poiCellStyle.setDefaultCellAlignment(hAlign, vAlign, wrapText);
		
		//protect
		boolean locked = cellStyle.isLocked();
		boolean hidden = cellStyle.isHidden();
		poiCellStyle.setDefaultProtection(locked, hidden);

		return poiCellStyle;
	}

	//ZSS-855
	@Override
	protected int exportTables(SSheet sheet, Sheet poiSheet0, int tbId) {
		final XSSFSheet poiSheet = (XSSFSheet) poiSheet0;
		for (STable table : sheet.getTables()) {
			XSSFTable poiTable = poiSheet.createTable();
			poiTable.setName(table.getName());
			poiTable.setDisplayName(table.getDisplayName());
			poiTable.setRef(table.getAllRegion().getRegion().getReferenceString());
			poiTable.setTotalsRowCount(table.getTotalsRowCount());
			poiTable.setHeaderRowCount(table.getHeaderRowCount());
			final XSSFTableStyleInfo poiInfo = poiTable.createTableStyleInfo();
			final STableStyleInfo info = table.getTableStyleInfo();
			poiInfo.setName(info.getName());
			poiInfo.setShowColumnStripes(info.isShowColumnStripes());
			poiInfo.setShowRowStripes(info.isShowRowStripes());
			poiInfo.setShowLastColumn(info.isShowLastColumn());
			poiInfo.setShowFirstColumn(info.isShowFirstColumn());
			
			final SAutoFilter filter = table.getAutoFilter();
			if (filter != null) {
				final CellRegion region = filter.getRegion();
				XSSFAutoFilter poiFilter = poiTable.createAutoFilter();
				poiFilter.setRef(region.getReferenceString());
				exportFilterColumns(poiFilter, filter, region.getColumnCount());
			} else {
				poiTable.clearAutoFilter();
			}
			
			int j = 0;
			for (STableColumn tbCol : table.getColumns()) {
				final XSSFTableColumn poiTbCol = poiTable.addTableColumn();
				poiTbCol.setName(tbCol.getName());
				poiTbCol.setId(++j);
				if (tbCol.getTotalsRowFunction() != null)
					poiTbCol.setTotalsRowFunction(TotalsRowFunction.values()[tbCol.getTotalsRowFunction().ordinal()]);
				if (tbCol.getTotalsRowFunction() == STotalsRowFunction.none && tbCol.getTotalsRowLabel() != null)
					poiTbCol.setTotalsRowLabel(tbCol.getTotalsRowLabel());
				else if (tbCol.getTotalsRowFunction() == STotalsRowFunction.custom && tbCol.getTotalsRowFormula() != null)
					poiTbCol.setTotalsRowFormula(tbCol.getTotalsRowFormula()); //ZSS-977
			}
			poiTable.setId(++tbId);
			((XSSFWorkbook)workbook).addTableName(poiTable);
		}
		
		return tbId;
	}
	
	//ZSS-1145
	protected void addPOIDxfCellStyle(SExtraStyle extraStyle) {
		// instead of creating a new style, use old one if exist
		XSSFDxfCellStyle poiCellStyle = (XSSFDxfCellStyle) styleTable.get(extraStyle);
		if (poiCellStyle != null) {
//			workbook.addDxfCellStyle(poiCellStyle); //ZSS-1191: already in workbook; no need to add again
			return;
		}
		poiCellStyle = (XSSFDxfCellStyle) workbook.createDxfCellStyle(); //will add into workbook

		// Border
		if (extraStyle.getBorder() != null) {
			final BorderType btb = extraStyle.getBorderBottom();
			final BorderType btl = extraStyle.getBorderLeft();
			final BorderType btr = extraStyle.getBorderRight();
			final BorderType btt = extraStyle.getBorderTop();
			final BorderType btd = extraStyle.getBorderDiagonal();
			final BorderType bth = extraStyle.getBorderHorizontal();
			final BorderType btv = extraStyle.getBorderVertical();
			
			final short bottom = btb == null ? -1 : PoiEnumConversion.toPoiBorderType(btb);
			final short left = btl == null ? -1 : PoiEnumConversion.toPoiBorderType(btl);
			final short right = btr == null ? -1 : PoiEnumConversion.toPoiBorderType(btr);
			final short top =  btt == null ? -1 : PoiEnumConversion.toPoiBorderType(btt);
			final short diagonal = btd == null ? -1 : PoiEnumConversion.toPoiBorderType(btd);
			final short horizontal = bth == null ? -1 : PoiEnumConversion.toPoiBorderType(bth);
			final short vertical = btv == null ? -1 : PoiEnumConversion.toPoiBorderType(btv); 
			
			final Color bottomColor = toPOIColor(extraStyle.getBorderBottomColor());
			final Color leftColor = toPOIColor(extraStyle.getBorderLeftColor());
			final Color rightColor = toPOIColor(extraStyle.getBorderRightColor());
			final Color topColor = toPOIColor(extraStyle.getBorderTopColor());
			final Color diagonalColor = toPOIColor(extraStyle.getBorderDiagonalColor());
			final Color horizontalColor = toPOIColor(extraStyle.getBorderHorizontalColor());
			final Color verticalColor = toPOIColor(extraStyle.getBorderVerticalColor());
			
			boolean diaUp = extraStyle.isShowDiagonalUpBorder();
			boolean diaDn = extraStyle.isShowDiagonalDownBorder();
			
			poiCellStyle.setBorder(left, leftColor, top, topColor, right, rightColor, bottom, bottomColor, diagonal, diagonalColor, horizontal, horizontalColor, vertical, verticalColor, diaUp, diaDn);
		}

		// Fill
		final AbstractFillAdv fill = (AbstractFillAdv) extraStyle.getFill();
		if (fill != null) {
			SColor fgColor = fill.getRawFillColor();
			SColor bgColor = fill.getRawBackColor();
			FillPattern fillPattern = fill.getRawFillPattern();

// ZSS-992: in <dxf>, bgColor and fgColor is set as is; (whilst <xf> is reversed)		
//			if (fillPattern == null || fillPattern == FillPattern.SOLID) { //ZSS-1162
//				SColor tmp = fgColor;
//				fgColor = bgColor;
//				bgColor = tmp;
//			}
			Color fillColor = fgColor == null ? null : toPOIColor(fgColor);
			Color backColor = bgColor == null ? null : toPOIColor(bgColor);
			short pattern = fillPattern == null ? -1 : PoiEnumConversion.toPoiFillPattern(extraStyle.getFillPattern());
			poiCellStyle.setFill(fillColor, backColor, pattern);
		}
		
//		//cell Alignment
//		short hAlign = PoiEnumConversion.toPoiHorizontalAlignment(extraStyle.getAlignment());
//		short vAlign = PoiEnumConversion.toPoiVerticalAlignment(extraStyle.getVerticalAlignment());
//		boolean wrapText = extraStyle.isWrapText();
//		
//		//ZSS-1020
//		poiCellStyle.setCellAlignment(hAlign, vAlign, wrapText, (short) extraStyle.getRotation());
//		
//		//protect
//		boolean locked = extraStyle.isLocked();
//		boolean hidden = extraStyle.isHidden();
//		poiCellStyle.setProtection(locked, hidden);

		// NumFmt
		// TODO: if custom formatCode?
		final String dataFormat = extraStyle.getDataFormat(); 
		if (dataFormat != null) {
			DataFormat df = workbook.createDataFormat();
			short fmt = df.getFormat(dataFormat);
			poiCellStyle.setDataFormat(fmt);
		}

		// Font
		poiCellStyle.setFont(toPOIDxfFont(extraStyle.getFont()));

		// put into table
		styleTable.put(extraStyle, poiCellStyle);
		
//		int indention = extraStyle.getIndention();
//		if (indention > 0) 
//			poiCellStyle.setIndention((short) indention);
	}
	
	//ZSS-1145
	protected Font toPOIDxfFont(SFont font0) {
		if (font0 == null) return null; //ZSS-1138

		Font poiFont = XSSFFont.createDxfFont();
		AbstractFontAdv font = (AbstractFontAdv) font0;
		
		if (font.isOverrideBold())
			poiFont.setBoldweight(PoiEnumConversion.toPoiBoldweight(font.getBoldweight()));
		if (font.isOverrideStrikeout())
			poiFont.setStrikeout(font.isStrikeout());
		if (font.isOverrideItalic())
			poiFont.setItalic(font.isItalic());
		if (font.isOverrideColor())
			BookHelper.setFontColor(workbook, poiFont, toPOIColor(font.getColor()));
		if (font.isOverrideHeightPoints())
			poiFont.setFontHeightInPoints((short) font.getHeightPoints());
		if (font.isOverrideName())
			poiFont.setFontName(font.getName());
		if (font.isOverrideTypeOffset())
			poiFont.setTypeOffset(PoiEnumConversion.toPoiTypeOffset(font.getTypeOffset()));
		if (font.isOverrideUnderline())
			poiFont.setUnderline(PoiEnumConversion.toPoiUnderline(font.getUnderline()));

		return poiFont;
	}

	//ZSS-1141
	@Override
	protected void exportConditionalFormatting(SSheet sheet, Sheet poiSheet) {
		final List<SConditionalFormatting> formattings = sheet.getConditionalFormattings();
		int priority = formattings.size();
		for (SConditionalFormatting cf : formattings) {
			final XSSFConditionalFormatting poicf = new XSSFConditionalFormatting((XSSFSheet) poiSheet);
			((XSSFSheet)poiSheet).addCondtionalFormatting(poicf);
			final CTConditionalFormatting ctcf = poicf.getCTConditionalFormatting();
			
			addSqref(ctcf, cf);
			
			for (SConditionalFormattingRule rule : cf.getRules()) {
				final CTCfRule ctRule = ctcf.addNewCfRule();
				addPoiRule(sheet, ctRule, rule, priority--);
			}
		}
	}
	
	//ZSS-1141
	protected void addSqref(CTConditionalFormatting ctcf, SConditionalFormatting cf) {
		StringBuilder sb = new StringBuilder();
		for (CellRegion rgn : cf.getRegions()) {
			if (sb.length() > 0) {
				sb.append(" ");
			}
			sb.append(rgn.getReferenceString());
		}
		List<String> sqrefs = new ArrayList<String>();
		sqrefs.add(sb.toString());
		ctcf.setSqref(sqrefs);
	}
	
	//ZSS-1141
	protected void addPoiRule(SSheet sheet, CTCfRule ctRule, SConditionalFormattingRule rule, int priority) {
		ctRule.setType(toConditionalFormattingRuleType(rule.getType()));
		ctRule.setPriority(priority);
		if (rule.isStopIfTrue()) {
			ctRule.setStopIfTrue(true);
		}
		final SExtraStyle extraStyle = rule.getExtraStyle();
		if (extraStyle != null) {
			addPOIDxfCellStyle(extraStyle);
			final XSSFDxfCellStyle poiCellStyle = (XSSFDxfCellStyle) styleTable.get(extraStyle);
			int index = poiCellStyle.getIndex();
			if (index >= 0)
				ctRule.setDxfId(index);
		}
		switch(rule.getType()) {
		case ABOVE_AVERAGE:
			if (!rule.isAboveAverage())
				ctRule.setAboveAverage(false);
			if (rule.isEqualAverage())
				ctRule.setEqualAverage(true);
			if (rule.getStandardDeviation() != null)
				ctRule.setStdDev(rule.getStandardDeviation());
			break;
		case CELL_IS:
			if (rule.getOperator() != null)
				ctRule.setOperator(toCFRuleOperator(rule.getOperator()));
			addFormulas(ctRule, rule);
			break;
		case COLOR_SCALE:
			if (rule.getColorScale() != null)
				addColorScale(ctRule, rule);
			break;
		case CONTAINS_BLANKS:
		case NOT_CONTAINS_BLANKS:
			addFormulas(ctRule, rule);
			break;
		case CONTAINS_ERRORS:
		case NOT_CONTAINS_ERRORS:
			addFormulas(ctRule, rule);
			break;
		case CONTAINS_TEXT:
		case NOT_CONTAINS_TEXT:
			if (rule.getText() != null)
				ctRule.setText(rule.getText());
			if (rule.getOperator() != null)
				ctRule.setOperator(toCFRuleOperator(rule.getOperator()));
			addFormulas(ctRule, rule);
			break;
		case DATA_BAR:
			if (rule.getDataBar() != null)
				addDataBar(ctRule, rule);
			break;
		case BEGINS_WITH:
		case ENDS_WITH:
			if (rule.getText() != null)
				ctRule.setText(rule.getText());
			if (rule.getOperator() != null)
				ctRule.setOperator(toCFRuleOperator(rule.getOperator()));
			addFormulas(ctRule, rule);
			break;
		case EXPRESSION:
			addFormulas(ctRule, rule);
			break;
		case ICON_SET:
			if (rule.getIconSet() != null)
				addIconSet(ctRule, rule);
			break;
		case TIME_PERIOD:
			if (rule.getTimePeriod() != null)
				ctRule.setTimePeriod(toTimePeriod(rule.getTimePeriod()));
			addFormulas(ctRule, rule);
			break;
		case TOP_10:
			if (rule.getRank() != null)
				ctRule.setRank(rule.getRank());
			if (rule.isPercent())
				ctRule.setPercent(true);
			if (rule.isBottom())
				ctRule.setBottom(true);
			break;
		case DUPLICATE_VALUES:
		case UNIQUE_VALUES:
			// No extra setting
			break;
		}
	}
	
	//ZSS-1141
	protected void addIconSet(CTCfRule ctRule, SConditionalFormattingRule rule) {
		final CTIconSet ctIconSet = ctRule.addNewIconSet();
		final SIconSet iconSet = rule.getIconSet();
		for (SCFValueObject vo : iconSet.getCFValueObjects()) {
			final CTCfvo ctvo = ctIconSet.addNewCfvo();
			addValueObject(ctvo, vo);
		}
		
		ctIconSet.setIconSet(toIconSetType(iconSet.getType()));
		if (iconSet.isPercent())
			ctIconSet.setPercent(true);
		if (iconSet.isReverse())
			ctIconSet.setReverse(true);
		if (!iconSet.isShowValue())
			ctIconSet.setShowValue(false);
	}
	
	//ZSS-1141
	protected void addColorScale(CTCfRule ctRule, SConditionalFormattingRule rule) {
		final SColorScale colorScale = rule.getColorScale();
		final CTColorScale ctColorScale = ctRule.addNewColorScale();
		for (SCFValueObject vo : colorScale.getCFValueObjects()) {
			final CTCfvo ctvo = ctColorScale.addNewCfvo();
			addValueObject(ctvo, vo);
		}
		
		for (SColor color : colorScale.getColors()) {
			CTColor ctColor = ctColorScale.addNewColor();
			ctColor.setRgb(((AbstractColorAdv)color).getARGB());
		}
	}
	
	//ZSS-1141
	protected void addDataBar(CTCfRule ctRule, SConditionalFormattingRule rule) {
		final CTDataBar ctDataBar = ctRule.addNewDataBar();
		final SDataBar dataBar = rule.getDataBar();
		for (SCFValueObject vo : dataBar.getCFValueObjects()) {
			final CTCfvo ctvo = ctDataBar.addNewCfvo();
			addValueObject(ctvo, vo);
		}
		CTColor ctColor = ctDataBar.addNewColor();
		ctColor.setRgb(((AbstractColorAdv)dataBar.getColor()).getARGB());
		
		if (dataBar.getMaxLength() != 90) {
			ctDataBar.setMaxLength(dataBar.getMaxLength());
		}
		if (dataBar.getMinLength() != 10) {
			ctDataBar.setMinLength(dataBar.getMinLength());
		}
		if (!dataBar.isShowValue()) {
			ctDataBar.setShowValue(false);
		}
	}

	//ZSS-1141
	protected void addFormulas(CTCfRule ctRule, SConditionalFormattingRule rule) {
		//ZSS-1142
		if (rule.getFormula1() != null) {
			ctRule.addFormula(rule.getFormula1().substring(1)); // don't include the leading = character
		}
		if (rule.getFormula2() != null) {
			ctRule.addFormula(rule.getFormula2().substring(1)); // don't include the leading = character
		}
		if (rule.getFormula3() != null) {
			ctRule.addFormula(rule.getFormula3().substring(1)); // don't include the leading = character
		}
	}
	
	//ZSS-1141
	protected void addValueObject(CTCfvo ctvo, SCFValueObject vo) {
		if (vo.isGreaterOrEqual()) {
			//ZSS-1142
			if (ctvo.isSetGte()) {
				ctvo.unsetGte();
			}
		}
		ctvo.setType(toValueObjectType(vo.getType()));
		if (vo.getValue() != null) {
			ctvo.setVal(vo.getValue());
		}
	}
	
	//ZSS-1141
	protected STTimePeriod.Enum toTimePeriod(SConditionalFormattingRule.RuleTimePeriod ctPeriod) {
		return STTimePeriod.Enum.forInt(ctPeriod.value);
	}
	//ZSS-1141
	protected STConditionalFormattingOperator.Enum toCFRuleOperator(SConditionalFormattingRule.RuleOperator ctType) {
		return STConditionalFormattingOperator.Enum.forInt(ctType.value);
	}
	
	//ZSS-1141
	protected STIconSetType.Enum toIconSetType(SIconSet.IconSetType  ctType) {
		return STIconSetType.Enum.forInt(ctType.value);
	}
	
	//ZSS-1141
	protected STCfvoType.Enum toValueObjectType(SCFValueObject.CFValueObjectType ctType) {
		return STCfvoType.Enum .forInt(ctType.value);
	}
	
	//ZSS-1141
	protected STCfType.Enum toConditionalFormattingRuleType(SConditionalFormattingRule.RuleType cfType) {
		return STCfType.Enum.forInt(cfType.value);
	}
	
	//ZSS-1141
	protected STCfType.Enum toConditionalFormatingRuleType(SConditionalFormattingRule.RuleType stype) {
		return STCfType.Enum.forInt(stype.value);
	}

	//ZSS-1188
	protected void addPOITableStyle(STableStyle tbStyle) {
		// instead of creating a new style, use old one if exist
		XSSFTableStyle poiTbStyle = (XSSFTableStyle) tbStyleTable.get(tbStyle);
		if (poiTbStyle != null) {
			workbook.addTableStyle(poiTbStyle);
			return;
		}
		poiTbStyle = (XSSFTableStyle) workbook.createTableStyle(tbStyle.getName()); //will add into workbook

		final STableStyleElem wholeTable = tbStyle.getWholeTableStyle();
		if (wholeTable != null) {
			final int dxfId = getOrCreateDxfId(wholeTable);
			poiTbStyle.addTableStyleElement(dxfId, "wholeTable");
		}
		
		final STableStyleElem colStripe1 = tbStyle.getColStripe1Style();
		if (colStripe1 != null) {
			final int colStripe1Size = tbStyle.getColStripe1Size();
			final int dxfId = getOrCreateDxfId(colStripe1);
			poiTbStyle.addTableStyleElement(dxfId, "firstColumnStripe", colStripe1Size);
		}
		
		final STableStyleElem colStripe2 = tbStyle.getColStripe2Style();
		if (colStripe2 != null) {
			final int colStripe2Size = tbStyle.getColStripe2Size();
			final int dxfId = getOrCreateDxfId(colStripe2);
			poiTbStyle.addTableStyleElement(dxfId, "SecondColumnStripe", colStripe2Size);
		}
		
		final STableStyleElem rowStripe1 = tbStyle.getRowStripe1Style();
		if (rowStripe1 != null) {
			final int rowStripe1Size = tbStyle.getRowStripe1Size();
			final int dxfId = getOrCreateDxfId(rowStripe1);
			poiTbStyle.addTableStyleElement(dxfId, "firstRowStripe", rowStripe1Size);
		}
		
		final STableStyleElem rowStripe2 = tbStyle.getRowStripe2Style();
		if (rowStripe2 != null) {
			final int rowStripe2Size = tbStyle.getRowStripe2Size();
			final int dxfId = getOrCreateDxfId(rowStripe2);
			poiTbStyle.addTableStyleElement(dxfId, "firstColumnStripe", rowStripe2Size);
		}
		
		final STableStyleElem lastColumn = tbStyle.getLastColumnStyle();
		if (lastColumn != null) {
			final int dxfId = getOrCreateDxfId(lastColumn);
			poiTbStyle.addTableStyleElement(dxfId, "lastColumn");
		}
			
		final STableStyleElem firstColumn = tbStyle.getFirstColumnStyle();
		if (firstColumn != null) {
			final int dxfId = getOrCreateDxfId(firstColumn);
			poiTbStyle.addTableStyleElement(dxfId, "firstColumn");
		}
		
		final STableStyleElem headerRow = tbStyle.getHeaderRowStyle();
		if (headerRow != null) {
			final int dxfId = getOrCreateDxfId(headerRow);
			poiTbStyle.addTableStyleElement(dxfId, "headerRow");
		}
		
		final STableStyleElem totalRow = tbStyle.getTotalRowStyle();
		if (totalRow != null) {
			final int dxfId = getOrCreateDxfId(totalRow);
			poiTbStyle.addTableStyleElement(dxfId, "totalRow");
		}
		
		final STableStyleElem firstHeaderCell = tbStyle.getFirstHeaderCellStyle();
		if (firstHeaderCell != null) {
			final int dxfId = getOrCreateDxfId(firstHeaderCell);
			poiTbStyle.addTableStyleElement(dxfId, "firstHeaderCell");
		}
		
		final STableStyleElem lastHeaderCell = tbStyle.getLastHeaderCellStyle();
		if (lastHeaderCell != null) {
			final int dxfId = getOrCreateDxfId(lastHeaderCell);
			poiTbStyle.addTableStyleElement(dxfId, "lastHeaderCell");
		}

		final STableStyleElem firstTotalCell = tbStyle.getFirstTotalCellStyle();
		if (firstTotalCell != null) {
			final int dxfId = getOrCreateDxfId(firstTotalCell);
			poiTbStyle.addTableStyleElement(dxfId, "firstTotalCell");
		}
		
		final STableStyleElem lastTotalCell = tbStyle.getLastTotalCellStyle();
		if (lastTotalCell != null) {
			final int dxfId = getOrCreateDxfId(lastTotalCell);
			poiTbStyle.addTableStyleElement(dxfId, "lastTotalCell");
		}

		// put into table
		tbStyleTable.put(tbStyle, poiTbStyle);
	}
	
	//ZSS-1188
	protected int getOrCreateDxfId(STableStyleElem tbStyleElem) {
		int index = getOrCreateDxfId0(tbStyleElem);
		if (index < 0) {
			addPOIDxfCellStyle(tbStyleElem);
			return getOrCreateDxfId0(tbStyleElem);
		}
		return index;
	}
	//ZSS-1188
	private int getOrCreateDxfId0(STableStyleElem tbStyleElem) {
		XSSFDxfCellStyle poiCellStyle = (XSSFDxfCellStyle) styleTable.get(tbStyleElem);
		if (poiCellStyle != null) {
			return workbook.getDxfIndex(poiCellStyle);
		}
		return -1;
	}
	//ZSS-1189
	@Override
	protected RichTextString toPOIRichText(SRichText richText) {
		//ZSS-1189
		CreationHelper helper = workbook.getCreationHelper();
		XSSFRichTextString poiRichTextString = 
				(XSSFRichTextString) helper.createRichTextString(richText.getText());

		for (Segment sg : richText.getSegments()) {
			SFont font = sg.getFont();
			String text = sg.getText();
			poiRichTextString.addRun(text, (XSSFFont)toPOIFont(font));
		}

		return poiRichTextString;
	}

}
