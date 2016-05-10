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

import java.io.*;
import java.util.*;
import java.util.logging.*;

import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.w3c.dom.Node;
import org.zkoss.poi.POIXMLDocumentPart;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.usermodel.charts.*;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.model.ExternalLink;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.poi.xssf.usermodel.charts.*;
import org.zkoss.zss.model.*;
import org.zkoss.zss.model.SAutoFilter.SColorFilter;
import org.zkoss.zss.model.SBorder.BorderType;
import org.zkoss.zss.model.SChart.ChartType;
import org.zkoss.zss.model.SConditionalFormattingRule.RuleType;
import org.zkoss.zss.model.SFill.FillPattern;
import org.zkoss.zss.model.STableColumn.STotalsRowFunction;
import org.zkoss.zss.model.chart.*;
import org.zkoss.zss.model.impl.AbstractBorderAdv;
import org.zkoss.zss.model.impl.AbstractDataValidationAdv;
import org.zkoss.zss.model.impl.AbstractFillAdv;
import org.zkoss.zss.model.impl.AbstractFontAdv;
import org.zkoss.zss.model.impl.AbstractSheetAdv;
import org.zkoss.zss.model.impl.BorderImpl;
import org.zkoss.zss.model.impl.BorderLineImpl;
import org.zkoss.zss.model.impl.CFValueObjectImpl;
import org.zkoss.zss.model.impl.ChartAxisImpl;
import org.zkoss.zss.model.impl.ColorFilterImpl;
import org.zkoss.zss.model.impl.ColorImpl;
import org.zkoss.zss.model.impl.ColorScaleImpl;
import org.zkoss.zss.model.impl.ConditionalFormattingRuleImpl;
import org.zkoss.zss.model.impl.DataBarImpl;
import org.zkoss.zss.model.impl.ExtraFillImpl;
import org.zkoss.zss.model.impl.ExtraStyleImpl;
import org.zkoss.zss.model.impl.FillImpl;
import org.zkoss.zss.model.impl.IconSetImpl;
import org.zkoss.zss.model.impl.TableColumnImpl;
import org.zkoss.zss.model.impl.TableImpl;
import org.zkoss.zss.model.impl.TableStyleImpl;
import org.zkoss.zss.model.impl.TableStyleInfoImpl;
import org.zkoss.zss.model.impl.TableStyleElemImpl;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.util.CellStyleMatcher;
import org.zkoss.zss.model.impl.AbstractBookAdv;
import org.zkoss.zss.model.impl.ConditionalFormattingImpl;

/**
 * Specific importing behavior for XLSX.
 * 
 * @author Hawk
 * @since 3.5.0
 */
public class ExcelXlsxImporter extends AbstractExcelImporter {
	private static final long serialVersionUID = -1531963712987770117L;
	
	private static final Logger logger = Logger.getLogger(ExcelXlsxImporter.class.getName());

	@Override
	protected Workbook createPoiBook(InputStream is) throws IOException {
		return new XSSFWorkbook(is);
	}
	
	@Override
	protected void importExternalBookLinks() {
		try {
			// directly use XML to retrieve external book references and fetch external book names
			// if retrieving ExternalLink from relations list of workbook, the order might be incorrect.
			// according to spec. ISO 29500-1, 18.17.2.3, p. 2048, the order should be natural order of XML tags.
			List<String> bookNames = new ArrayList<String>();
			XSSFWorkbook xssfBook = (XSSFWorkbook)workbook;
			CTWorkbook ctBook = xssfBook.getCTWorkbook();
			CTExternalReferences externalReferences = ctBook.getExternalReferences();
			if(externalReferences != null) {
				List<CTExternalReference> exRefs = externalReferences.getExternalReferenceList();
				for(CTExternalReference exRef : exRefs) {
					ExternalLink link = (ExternalLink)xssfBook.getRelationById(exRef.getId());
					bookNames.add(link.getBookName());
				}
			}
			if(bookNames.size() > 0) {
				this.book.setAttribute(FormulaEngine.KEY_EXTERNAL_BOOK_NAMES, bookNames);
			}
		} catch(Exception e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
		}
	}


	/**
	 * [ISO/IEC 29500-1 1st Edition] 18.3.1.13 col (Column Width & Formatting)
	 * By experiments, CT_Col is always created in ascending order by min and the range specified by min & max doesn't overlap each other.<br/>
	 * For example:
	 * <x:cols xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	 *   <x:col min="1" max="1" width="11.28515625" customWidth="1" />
	 *   <x:col min="2" max="4" width="11.28515625" hidden="1" customWidth="1" />
	 *   <x:col min="5" max="5" width="11.28515625" customWidth="1" />
	 * </x:cols> 
	 */
	@Override
	protected void importColumn(Sheet poiSheet, SSheet sheet) {
		CTWorksheet worksheet = ((XSSFSheet)poiSheet).getCTWorksheet();
		if(worksheet.sizeOfColsArray()<=0){
			return;
		}

		//ZSS-1132
		final int charWidth = ((AbstractBookAdv)book).getCharWidth();
		
		//ZSS-952: xssf's default column width includes the 5 pixels padding and is in charwidth unit.
//		int defaultWidth = UnitUtil.defaultColumnWidthToPx(poiSheet.getDefaultColumnWidth(), CHRACTER_WIDTH);
		int defaultWidth = UnitUtil.xssfDefaultColumnWidthToPx(((XSSFSheet)poiSheet).getXssfDefaultColumnWidth(), charWidth);
		CTCols colsArray = worksheet.getColsArray(0);
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
			CTCol ctCol = colsArray.getColArray(i);
			//max is 16384
			
			SColumnArray columnArray = sheet.setupColumnArray((int)ctCol.getMin()-1, (int)ctCol.getMax()-1);
			
			boolean hidden = ctCol.getHidden();
			int columnIndex = (int)ctCol.getMin()-1;
			columnArray.setHidden(hidden);
			int width = ImExpUtils.getWidthAny(poiSheet, columnIndex, charWidth);
			if (!(hidden || width == defaultWidth)){
				//when CT_Col is hidden with default width, We don't import the width for it's 0.  
				columnArray.setWidth(width);
			}
			columnArray.setCustomWidth(ctCol.getCustomWidth() || width != defaultWidth); //ZSS-896

			CellStyle columnStyle = poiSheet.getColumnStyle(columnIndex);
			if (columnStyle != null){
				columnArray.setCellStyle(importCellStyle(columnStyle));
			}
		}
	}

	/**
	 * Not import X & Y axis title because {@link XSSFCategoryAxis} doesn't provide API to get title. 
	 * Reference ChartHelper.drawXSSFChart()
	 */
	protected void importChart(List<ZssChartX> poiCharts, Sheet poiSheet, SSheet sheet) {
		
		for (ZssChartX zssChart : poiCharts){
			XSSFChart xssfChart = (XSSFChart)zssChart.getChart();
			ViewAnchor viewAnchor = toViewAnchor(poiSheet, xssfChart.getPreferredSize());

			SChart chart = null;
			CategoryData categoryData = null;
			//reference ChartHelper.drawXSSFChart()
			switch(xssfChart.getChartType()){
				case Area:
					chart = sheet.addChart(ChartType.AREA, viewAnchor);
					categoryData = new XSSFAreaChartData(xssfChart);
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFAreaChartData)categoryData).getGrouping()));
					break;
				case Area3D:
					chart = sheet.addChart(ChartType.AREA, viewAnchor);
					categoryData = new XSSFArea3DChartData(xssfChart);
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFArea3DChartData)categoryData).getGrouping()));
					break;
				case Bar:
					chart = sheet.addChart(ChartType.BAR, viewAnchor);
					categoryData = new XSSFBarChartData(xssfChart);
					chart.setBarDirection(PoiEnumConversion.toBarDirection(((XSSFBarChartData)categoryData).getBarDirection()));
					chart.setBarOverlap(xssfChart.getBarOverlap()); //ZSS-830
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFBarChartData)categoryData).getGrouping()));
					break;
				case Bar3D:
					chart = sheet.addChart(ChartType.BAR, viewAnchor);
					categoryData = new XSSFBar3DChartData(xssfChart);
					chart.setBarDirection(PoiEnumConversion.toBarDirection(((XSSFBar3DChartData)categoryData).getBarDirection()));
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFBar3DChartData)categoryData).getGrouping()));
					break;
				case Bubble:
					chart = sheet.addChart(ChartType.BUBBLE, viewAnchor);
					
					XYZData xyzData = new XSSFBubbleChartData(xssfChart);
					importXyzSeries(xyzData.getSeries(), (SGeneralChartData)chart.getData());
					break;
				case Column:
					chart = sheet.addChart(ChartType.COLUMN, viewAnchor);
					categoryData = new XSSFColumnChartData(xssfChart);
					chart.setBarDirection(PoiEnumConversion.toBarDirection(((XSSFColumnChartData)categoryData).getBarDirection()));
					chart.setBarOverlap(xssfChart.getBarOverlap()); //ZSS-830
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFColumnChartData)categoryData).getGrouping()));
					break;
				case Column3D:
					chart = sheet.addChart(ChartType.COLUMN, viewAnchor);
					categoryData = new XSSFColumn3DChartData(xssfChart);
					chart.setBarDirection(PoiEnumConversion.toBarDirection(((XSSFColumn3DChartData)categoryData).getBarDirection()));
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFColumn3DChartData)categoryData).getGrouping()));
					break;
				case Doughnut:
					chart = sheet.addChart(ChartType.DOUGHNUT, viewAnchor);
					categoryData = new XSSFDoughnutChartData(xssfChart);
					break;
				case Line:
					chart = sheet.addChart(ChartType.LINE, viewAnchor);
					categoryData = new XSSFLineChartData(xssfChart);
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFLineChartData)categoryData).getGrouping())); //ZSS-828
					break;
				case Line3D:
					chart = sheet.addChart(ChartType.LINE, viewAnchor);
					categoryData = new XSSFLine3DChartData(xssfChart);
					chart.setGrouping(PoiEnumConversion.toChartGrouping(((XSSFLine3DChartData)categoryData).getGrouping())); //ZSS-828
					break;
				case Pie:
					chart = sheet.addChart(ChartType.PIE, viewAnchor);
					categoryData = new XSSFPieChartData(xssfChart);
					break;
				case Pie3D:
					chart = sheet.addChart(ChartType.PIE, viewAnchor);
					categoryData = new XSSFPie3DChartData(xssfChart);
					break;
				case Scatter:
					chart = sheet.addChart(ChartType.SCATTER, viewAnchor);
					
					XYData xyData =  new XSSFScatChartData(xssfChart);
					importXySeries(xyData.getSeries(), (SGeneralChartData)chart.getData());
					break;
//				case Stock:
//					chart = sheet.addChart(NChartType.STOCK, viewAnchor);
//					categoryData = new XSSFStockChartData(xssfChart); TODO XSSFStockChartData contains wrong implementation
//					break;
				default:
					//TODO ignore unsupported charts
					continue;
			}
			
			if (xssfChart.getTitle()!=null){
				chart.setTitle(xssfChart.getTitle().getString());
			}
			chart.setThreeD(xssfChart.isSetView3D());
			//ZSS-830
			if (chart.isThreeD()) {
				XSSFView3D view3d = xssfChart.getOrCreateView3D();
				chart.setRotX(view3d.getRotX());
				chart.setRotY(view3d.getRotY());
				chart.setPerspective(view3d.getPerspective());
				chart.setHPercent(view3d.getHPercent());
				chart.setDepthPercent(view3d.getDepthPercent());
				chart.setRightAngleAxes(view3d.isRightAngleAxes());
			}
			if (xssfChart.hasLegend()) {
				chart.setLegendPosition(PoiEnumConversion.toLengendPosition(xssfChart.getOrCreateLegend().getPosition()));
			}
			if (categoryData != null){
				importSeries(categoryData.getSeries(), (SGeneralChartData)chart.getData());
			}
			//ZSS-822
			importAxis(xssfChart, chart);
			
			chart.setPlotOnlyVisibleCells(xssfChart.isPlotOnlyVisibleCells());
		}
	}
	//ZSS-822
	protected void importAxis(XSSFChart xssfChart, SChart chart) {
		List axises = (List) xssfChart.getAxis();
		if (axises != null) {
			for (Object axis0 : axises) {
				if (axis0 instanceof ValueAxis) {
					ValueAxis axis = (ValueAxis) axis0;
					String format = ((ValueAxis) axis).getNumberFormat();
					double min = axis.getMinimum();
					double max = axis.getMaximum();
					SChartAxis saxis = new ChartAxisImpl(axis.getId(), SChartAxis.SChartAxisType.VALUE, min, max, format);
					chart.addValueAxis(saxis);
				} else if (axis0 instanceof CategoryAxis) {
					CategoryAxis axis = (CategoryAxis) axis0;
					String format = null;
					double min = axis.getMinimum();
					double max = axis.getMaximum();
					SChartAxis saxis = new ChartAxisImpl(axis.getId(), SChartAxis.SChartAxisType.CATEGORY, min, max, format);
					chart.addCategoryAxis(saxis);
				}
			}
		}
	}

	/**
	 * reference ChartHepler.prepareCategoryModel() 
	 * Category normally indicates the values show on X axis.
	 * Tile indicates the name for each series.
	 * @param seriesList source chart data
	 * @param chartData destination chart data
	 */
	protected void importSeries(List<? extends CategoryDataSerie> seriesList, SGeneralChartData chartData) {
		CategoryDataSerie firstSeries = null;
		if ((firstSeries = seriesList.get(0))!=null){
			chartData.setCategoriesFormula(getValueFormula(firstSeries.getCategories()));
		}
		for (int i =0 ;  i< seriesList.size() ; i++){
			CategoryDataSerie sourceSeries = seriesList.get(i);
			String nameExpression = getTitleFormula(sourceSeries.getTitle(), i);			
			String xValueExpression = getValueFormula(sourceSeries.getValues());
			SSeries series = chartData.addSeries();
			series.setFormula(nameExpression, xValueExpression);
		}
	}
	
	protected void importXySeries(List<? extends XYDataSerie> seriesList, SGeneralChartData chartData) {
		for (int i =0 ;  i< seriesList.size() ; i++){
			XYDataSerie sourceSeries = seriesList.get(i);
			SSeries series = chartData.addSeries();
			series.setXYFormula(getTitleFormula(sourceSeries.getTitle(), i), 
								getValueFormula(sourceSeries.getXs()), 
								getValueFormula(sourceSeries.getYs()));
		}
	}

	/**
	 * reference ChartHepler.prepareXYZModel() 
	 * @param seriesList
	 * @param chart
	 */
	protected void importXyzSeries(List<? extends XYZDataSerie> seriesList, SGeneralChartData chartData) {
		for (int i =0 ;  i< seriesList.size() ; i++){
			XYZDataSerie sourceSeries = seriesList.get(i);
			//reference to ChartHelper.prepareTitle()
			String nameExpression = getTitleFormula(sourceSeries.getTitle(), i);			
			String xValueExpression = getValueFormula(sourceSeries.getXs());
			String yValueExpression = getValueFormula(sourceSeries.getYs());
			String zValueExpression = getValueFormula(sourceSeries.getZs());	
			
			SSeries series = chartData.addSeries();
			series.setXYZFormula(nameExpression, xValueExpression, yValueExpression, zValueExpression);
		}
	}
	
	/**
	 * return a formula or generate a default title ("Series[N]")if title doesn't exist.
	 * reference ChartHelper.prepareTitle()
	 * @param textSource
	 * @return
	 */
	protected String getTitleFormula(ChartTextSource textSource, int seriesIndex){
		if (textSource == null){
			return "\"Series"+seriesIndex+"\"";
		}else {
			if (textSource.isReference()){
				return textSource.getFormulaString();
			}else{
				return "\""+textSource.getTextString()+"\"";
			}
		}
	}
	
	/**
	 * reference ChartHelper.prepareValues()
	 * @param dataSource
	 * @return
	 */
	protected String getValueFormula(ChartDataSource<?> dataSource){
		if (dataSource.isReference()){
			return dataSource.getFormulaString();
		}else{
			int count = dataSource.getPointCount();
			if(count<=0){
				return null;
			}
			StringBuilder expression = new StringBuilder("{");
			for (int i = 0; i < count; ++i) {
				Object value = dataSource.getPointAt(i); //Number or String
				if (value == null){
					if (dataSource.isNumeric()){
						expression.append("0");
					}else{
						expression.append("\"\"");
					}
				}else{
					//ZSS-577 it is possible a string array
					if(dataSource.isNumeric()){
						expression.append(value.toString());
					}else{
						expression.append("\"").append(value).append("\"");
					}
				}
				if (i != count-1){
					expression.append(",");
				}
			}
			expression.append("}");
			return expression.toString();
		}
	}
	
	
	/*
	 * import drawings which include charts and pictures.
	 * reference DrawingManagerImpl.initXSSFDrawings()
	 */
	@Override
	protected void importDrawings(Sheet poiSheet, SSheet sheet) {
		/**
		 * A list of POI chart wrapper loaded during import.
		 */
		List<ZssChartX> poiCharts = new ArrayList<ZssChartX>();
		List<Picture> poiPictures = new ArrayList<Picture>();
		
		XSSFDrawing patriarch = null;
		for(POIXMLDocumentPart dr : ((XSSFSheet)poiSheet).getRelations()){
			if(dr instanceof XSSFDrawing){
				patriarch = (XSSFDrawing) dr;
				break;
			}
		}
		if (patriarch != null) {			
			final CTDrawing ctdrawing = patriarch.getCTDrawing();
			for (CTTwoCellAnchor anchor: ctdrawing.getTwoCellAnchorArray()) {
				final CTMarker from = anchor.getFrom();
				final CTMarker to = anchor.getTo();
				XSSFClientAnchor clientAnchor = null;
				if (from != null && to != null){
					clientAnchor = new XSSFClientAnchor((int)from.getColOff(), (int)from.getRowOff(), (int)to.getColOff(), (int)to.getRowOff(), 
							from.getCol(), from.getRow(), to.getCol(), to.getRow());
				}
				CTPicture ctPicture = anchor.getPic();
				if (ctPicture == null){ 
					final CTGraphicalObjectFrame gfrm = anchor.getGraphicFrame();
					if (gfrm != null) {
						final XSSFChartX chartX = createXSSFChartX(patriarch, gfrm , clientAnchor);
						poiCharts.add(chartX);
					}
				}else{ 
					poiPictures.add(XSSFPictureHelper.newXSSFPicture(patriarch, clientAnchor, ctPicture));
				}
			}
		}
		importChart(poiCharts, poiSheet, sheet);
		importPicture(poiPictures, poiSheet, sheet);
	}
	
	protected XSSFChartX createXSSFChartX(XSSFDrawing patriarch, CTGraphicalObjectFrame gfrm , XSSFClientAnchor xanchor) {
		//chart
		final String name = gfrm.getNvGraphicFramePr().getCNvPr().getName();
		final CTGraphicalObject gobj = gfrm.getGraphic();
		final CTGraphicalObjectData gdata = gobj.getGraphicData();
		String chartId = null;
		// ZSS-358: must skip other XML nodes such as TEXT_NODE
		Node child = gdata.getDomNode().getFirstChild();
		while(child != null) {
			if("chart".equals(child.getLocalName())) {
				chartId = child.getAttributes().getNamedItemNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id").getNodeValue();
				break;
			}
			child = child.getNextSibling();
		}
		final XSSFChartX chartX = new XSSFChartX(patriarch, xanchor, name, chartId);
		return chartX;
	}
	
	/**
	 * Reference DefaultBookWidgetLoader.getXSSFWidthInPx()
	 */
	@Override
	protected int getAnchorWidthInPx(ClientAnchor anchor, Sheet poiSheet) {
		//ZSS-1132
		final int charWidth = ((AbstractBookAdv)book).getCharWidth();
		
		final int firstColumn = anchor.getCol1();
	    final int firstColumnWidth = ImExpUtils.getWidthAny(poiSheet,firstColumn, charWidth);
	    int offsetInFirstColumn = UnitUtil.emuToPx(anchor.getDx1());

	    final int anchorWidthInFirstColumn = firstColumnWidth - offsetInFirstColumn;  
		int anchorWidthInLastColumn = UnitUtil.emuToPx(anchor.getDx2());
	    
		final int lastColumn = anchor.getCol2();
	    //in between
	    int width = firstColumn == lastColumn ? anchorWidthInLastColumn - offsetInFirstColumn : anchorWidthInFirstColumn + anchorWidthInLastColumn;
	    width = Math.abs(width); // just in case
	    
	    for (int j = firstColumn+1; j < lastColumn; ++j) {
	    	width += ImExpUtils.getWidthAny(poiSheet,j, charWidth);
	    }
	    
	    return width;
	}
	
	/**
	 * DefaultBookWidgetLoader.getXSSFHeightInPx()
	 */
	@Override
	protected int getAnchorHeightInPx(ClientAnchor anchor, Sheet poiSheet) {
	    final int firstRow = anchor.getRow1();
	    int offsetInFirstRow = UnitUtil.emuToPx(anchor.getDy1());
	    //first row
	    final int firstRowHeight = ImExpUtils.getHeightAny(poiSheet,firstRow);
		final int anchorHeightInFirstRow = firstRowHeight - offsetInFirstRow;  
	    
	    //last row
		final int lastRow = anchor.getRow2();
		int anchorHeightInLastRow = UnitUtil.emuToPx(anchor.getDy2());
		int height = lastRow == firstRow ? anchorHeightInLastRow - offsetInFirstRow : anchorHeightInFirstRow + anchorHeightInLastRow ;
		height = Math.abs(height); // just in case
	    
	    //add inter-row height
	    for (int row = firstRow+1; row < lastRow; ++row) {
	    	height += ImExpUtils.getHeightAny(poiSheet,row);
	    }
	    
	    return height;
	}


	@Override
	protected int getXoffsetInPixel(ClientAnchor clientAnchor, Sheet poiSheet) {
		return UnitUtil.emuToPx(clientAnchor.getDx1());
	}

	@Override
	protected int getYoffsetInPixel(ClientAnchor clientAnchor, Sheet poiSheet) {
		return UnitUtil.emuToPx(clientAnchor.getDy1());
	}

	/**
	 * Reference BookHelper.validate()
	 */
	@Override
	protected void importValidation(Sheet poiSheet, SSheet sheet) {
		for (DataValidation poiValidation : poiSheet.getDataValidations()){
			
			SDataValidation dataValidation = sheet.addDataValidation(null, null);
			CellRangeAddress[] cellRangeAddresses = poiValidation.getRegions().getCellRangeAddresses();
			DataValidationConstraint poiConstraint = poiValidation.getValidationConstraint();
			// getExplicitListValues() will be represented as formula1
			dataValidation.setOperatorType(PoiEnumConversion.toOperatorType(poiConstraint.getOperator()));
			dataValidation.setValidationType(PoiEnumConversion.toValidationType(poiConstraint.getValidationType()));
			((AbstractDataValidationAdv)dataValidation).setEscapedFormulas(poiConstraint.getFormula1(), poiConstraint.getFormula2());
			
			dataValidation.setIgnoreBlank(poiValidation.getEmptyCellAllowed());
			dataValidation.setErrorTitle(poiValidation.getErrorBoxTitle());
			dataValidation.setErrorMessage(poiValidation.getErrorBoxText());
			dataValidation.setAlertStyle(PoiEnumConversion.toErrorStyle(poiValidation.getErrorStyle()));
			dataValidation.setInputTitle(poiValidation.getPromptBoxTitle());
			dataValidation.setInputMessage(poiValidation.getPromptBoxText());
			if (poiConstraint.getValidationType() == DataValidationConstraint.ValidationType.LIST){
				/* 
				 * According to ISO/IEC 29500-1 \ 18.3.1.32  dataValidation (Data Validation) 
				 * Excel file contains reversed value against format specification . If not showing drop down, file contains 1 (true). If showing, 
				 * attribute "showDropDown" won't exist and POI get false. But POI's API reverse its value again.
				 * So we just directly accept the POI returned value.
				 */
				dataValidation.setInCellDropdown(poiValidation.getSuppressDropDownArrow());
			}
			dataValidation.setShowError(poiValidation.getShowErrorBox());
			dataValidation.setShowInput(poiValidation.getShowPromptBox());
			/*
			 * According to ISO/IEC 29500-1 \ 18.18.76  ST_Sqref (Reference Sequence) and A.2 
			 * Its XML Schema indicates it's a required attribute, so CellRangeAddresses must have at least one address. 
			 */
			Set<CellRegion> regions = new HashSet<CellRegion>();
			for(CellRangeAddress cellRangeAddr:cellRangeAddresses){
				regions.add(new CellRegion(cellRangeAddr.formatAsString()));
			}
			dataValidation.setRegions(regions);
		}
	}
	
	@Override
	protected boolean skipName(Name definedName) {
		boolean r = super.skipName(definedName);
		if(r)
			return r;
		if(((XSSFName)definedName).isBuiltInName()){
			return true;
		}
		return false;
	}
	
	@Override
	protected void setBookType(SBook book){
		book.setAttribute(BOOK_TYPE_KEY, "xlsx");
	}

	@Override
	protected void importPassword(Sheet poiSheet, SSheet sheet) {
		short hashpass = ((XSSFSheet)poiSheet).getPasswordHash(); 
		sheet.setHashedPassword(hashpass);
		
		//ZSS-1063
		String hashValue = ((XSSFSheet)poiSheet).getHashValue();
		sheet.setHashValue(hashValue);
		String spinCount = ((XSSFSheet)poiSheet).getSpinCount();
		sheet.setSpinCount(spinCount);
		String saltValue = ((XSSFSheet)poiSheet).getSaltValue();
		sheet.setSaltValue(saltValue);
		String algName = ((XSSFSheet)poiSheet).getAlgName();
		sheet.setAlgName(algName);
	}
	
	@Override
	protected void importSheetProtection(Sheet poiSheet, SSheet sheet) { //ZSS-576
		SheetProtection sp = poiSheet.getOrCreateSheetProtection();
		SSheetProtection ssp = sheet.getSheetProtection();
		
	    ssp.setAutoFilter(sp.isAutoFilter());
	    ssp.setDeleteColumns(sp.isDeleteColumns());
	    ssp.setDeleteRows(sp.isDeleteRows());
	    ssp.setFormatCells(sp.isFormatCells());
	    ssp.setFormatColumns(sp.isFormatColumns());
	    ssp.setFormatRows(sp.isFormatRows());
	    ssp.setInsertColumns(sp.isInsertColumns());
	    ssp.setInsertHyperlinks(sp.isInsertHyperlinks());
	    ssp.setInsertRows(sp.isInsertRows());
	    ssp.setPivotTables(sp.isPivotTables());
	    ssp.setSort(sp.isSort());
	    ssp.setObjects(sp.isObjects());
	    ssp.setScenarios(sp.isScenarios());
	    ssp.setSelectLockedCells(sp.isSelectLockedCells());
	    ssp.setSelectUnlockedCells(sp.isSelectUnlockedCells());
	}

	//ZSS-952
	@Override
	protected void importSheetDefaultColumnWidth(Sheet poiSheet, SSheet sheet) {
		//ZSS-1132
		final int charWidth = ((AbstractBookAdv)book).getCharWidth(); 

		// reference XUtils.getDefaultColumnWidthInPx()
		int defaultWidth = UnitUtil.xssfDefaultColumnWidthToPx(((XSSFSheet)poiSheet).getXssfDefaultColumnWidth(), charWidth);
		sheet.setDefaultColumnWidth(defaultWidth);
	}

	//ZSS-855
	@Override
	protected void importTables(Sheet poiSheet, SSheet sheet) {
		final XSSFSheet srcSheet = (XSSFSheet) poiSheet;
		for (XSSFTable poiTable : srcSheet.getTables()) {
			final CellReference cr1 = poiTable.getStartCellReference();
			final CellReference cr2 = poiTable.getEndCellReference();
			final CellRegion region0 = new CellRegion(cr1.getRow(), cr1.getCol(), cr2.getRow(), cr2.getCol());
			final SheetRegion region  = new SheetRegion(sheet, region0);
			final XSSFTableStyleInfo poiInfo = new XSSFTableStyleInfo(poiTable.getTableStyleInfo());
			final STableStyleInfo info = 
				new TableStyleInfoImpl(poiInfo.getName(), 
					poiInfo.isShowColumnStripes(), poiInfo.isShowRowStripes(),
					poiInfo.isShowFirstColumn(), poiInfo.isShowLastColumn());
 
			final STable table = new TableImpl((AbstractBookAdv)book,
				poiTable.getName(),
				poiTable.getDisplayName(),
				region,
				poiTable.getHeaderRowCount(),
				poiTable.getTotalsRowCount(),
				info);
			
			for (XSSFTableColumn poiTbCol : poiTable.getTableColumns()) {
				final STableColumn tbCol = new TableColumnImpl(poiTbCol.getName());
				tbCol.setTotalsRowFunction(STotalsRowFunction.values()[poiTbCol.getTotalsRowFunction().ordinal()]);
				if (tbCol.getTotalsRowFunction() == STotalsRowFunction.none)
					tbCol.setTotalsRowLabel(poiTbCol.getTotalsRowLabel());
				else if (tbCol.getTotalsRowFunction() == STotalsRowFunction.custom)
					tbCol.setTotalsRowFormula(poiTbCol.getTotalsRowFormula()); //ZSS-977
				table.addColumn(tbCol);
			}
			
			final XSSFAutoFilter poiFilter = poiTable.getAutoFilter();
			table.enableAutoFilter(poiFilter != null);
			
			sheet.addTable(table);

			//ZSS-1019
			if (poiFilter != null) {
				final int numberOfColumn = region.getColumnCount();
				final SAutoFilter zssFilter = table.getAutoFilter();
				importAutoFilterColumns(poiFilter, zssFilter, numberOfColumn);
			}
			
			importTableName(table);
			((AbstractBookAdv)book).addTable(table);
		}
	}
	
	//ZSS-855
	protected void importTableName(STable table) {
		SName namedRange = ((AbstractBookAdv)book).createTableName(table);
		SheetRegion rgn = table.getDataRegion();
		namedRange.setRefersToFormula(rgn.getReferenceString());
	}
	
	//ZSS-1140
	protected void importExtraStyles() {
		((AbstractBookAdv)book).clearExtraStyles();
		for (DxfCellStyle poiStyle : workbook.getDxfCellStyles()) {
			book.addExtraStyle(importExtraStyle(poiStyle));
		}
	}

	//ZSS-1140
	protected SExtraStyle importExtraStyle(DxfCellStyle poiCellStyle) {
		String dataFormat = poiCellStyle.getRawDataFormatString();
		FillPattern pattern = PoiEnumConversion.toFillPattern(poiCellStyle.getRawFillPattern());
		Color fgColor = poiCellStyle.getRawFillForegroundColor();
		Color bgColor = poiCellStyle.getRawFillBackgroundColor();
		SColor fgSColor = fgColor == null ? null : book.createColor(BookHelper.colorToForegroundHTML(workbook, fgColor));
		SColor bgSColor = bgColor == null ? null : book.createColor(BookHelper.colorToBackgroundHTML(workbook, bgColor));
// ZSS-992: in <dxf>, bgColor and fgColor is set as is; (whilst <xf> is reversed)		
//		if (pattern == null || pattern == FillPattern.SOLID) { //ZSS-1162
//			SColor tmp = fgSColor;
//			fgSColor = bgSColor;
//			bgSColor = tmp;
//		}
		SFill fill = new ExtraFillImpl(pattern, fgSColor, bgSColor); //ZSS-1162
		SBorder border = importBorder(poiCellStyle);
		SFont font = importFont(poiCellStyle);

		return new ExtraStyleImpl(font, fill, border, dataFormat);
	}

	//ZSS-1140
	protected SBorder importBorder(DxfCellStyle poiCellStyle) {
		if (!poiCellStyle.isOverrideBorder()) return null;
		
		BorderType typel = PoiEnumConversion.toBorderType(poiCellStyle.getRawBorderLeft());
		BorderType typet = PoiEnumConversion.toBorderType(poiCellStyle.getRawBorderTop());
		BorderType typer = PoiEnumConversion.toBorderType(poiCellStyle.getRawBorderRight());
		BorderType typeb = PoiEnumConversion.toBorderType(poiCellStyle.getRawBorderBottom());
		BorderType typed = PoiEnumConversion.toBorderType(poiCellStyle.getRawBorderDiagonal());
		BorderType typev = PoiEnumConversion.toBorderType(poiCellStyle.getRawBorderVertical());
		BorderType typeh = PoiEnumConversion.toBorderType(poiCellStyle.getRawBorderHorizontal());
		
		Color clrl = poiCellStyle.getRawBorderLeftColor();
		Color clrt = poiCellStyle.getRawBorderTopColor();
		Color clrr = poiCellStyle.getRawBorderRightColor();
		Color clrb = poiCellStyle.getRawBorderBottomColor();
		Color clrd = poiCellStyle.getRawBorderDiagonalColor();
		Color clrv = poiCellStyle.getRawBorderVerticalColor();
		Color clrh = poiCellStyle.getRawBorderHorizontalColor();
		
		SBorderLine left = typel == null ? null : new BorderLineImpl(typel, 
				clrl == null ? null : book.createColor(BookHelper.colorToBorderHTML(workbook, clrl)));
		SBorderLine top = typet == null ? null : new BorderLineImpl(typet,
				clrt == null ? null : book.createColor(BookHelper.colorToBorderHTML(workbook, clrt)));
		SBorderLine right = typer == null ? null : new BorderLineImpl(typer,
				clrr == null ? null : book.createColor(BookHelper.colorToBorderHTML(workbook, clrr)));
		SBorderLine bottom = typeb == null ? null : new BorderLineImpl(typeb,
				clrb == null ? null : book.createColor(BookHelper.colorToBorderHTML(workbook, clrb)));
		SBorderLine diagonal = typed == null ? null : new BorderLineImpl(typed, 
				clrd == null ? null : book.createColor(BookHelper.colorToBorderHTML(workbook, clrd)));
		SBorderLine vertical = typev == null ? null : new BorderLineImpl(typev, 
				clrv == null ? null : book.createColor(BookHelper.colorToBorderHTML(workbook, clrv)));
		SBorderLine horizontal = typeh == null ? null : new BorderLineImpl(typeh, 
				clrh == null ? null : book.createColor(BookHelper.colorToBorderHTML(workbook, clrh)));
		
		return new BorderImpl(left, top, right, bottom, diagonal, vertical, horizontal);
	}
	
	//ZSS-1140
	protected SFont createDxfZssFont(Font poiFont) {
		AbstractFontAdv font = (AbstractFontAdv) book.createFont(false);
		
		// font
		if (poiFont.isOverrideName())
			font.setName(poiFont.getFontName());
		if (poiFont.isOverrideBold())
			font.setBoldweight(PoiEnumConversion.toBoldweight(poiFont.getBoldweight()));
		if (poiFont.isOverrideItalic())
			font.setItalic(poiFont.getItalic());
		if (poiFont.isOverrideStrikeout())
			font.setStrikeout(poiFont.getStrikeout());
		if (poiFont.isOverrideUnderline())
			font.setUnderline(PoiEnumConversion.toUnderline(poiFont.getUnderline()));
		if (poiFont.isOverrideHeightPoints())
			font.setHeightPoints(poiFont.getFontHeightInPoints());
		if (poiFont.isOverrideTypeOffset())
			font.setTypeOffset(PoiEnumConversion.toTypeOffset(poiFont.getTypeOffset()));
		if (poiFont.isOverrideColor())
			font.setColor(book.createColor(BookHelper.getFontHTMLColor(workbook, poiFont)));
		
		font.setOverrideBold(poiFont.isOverrideBold());
		font.setOverrideColor(poiFont.isOverrideColor());
		font.setOverrideHeightPoints(poiFont.isOverrideHeightPoints());
		font.setOverrideItalic(poiFont.isOverrideItalic());
		font.setOverrideName(poiFont.isOverrideName());
		font.setOverrideStrikeout(poiFont.isOverrideStrikeout());
		font.setOverrideTypeOffset(poiFont.isOverrideTypeOffset());
		font.setOverrideUnderline(poiFont.isOverrideUnderline());
		return font;
	}

	//ZSS-1140
	@Override
	protected SFont importFont(CellStyle poiCellStyle) {
		if (poiCellStyle instanceof DxfCellStyle) {
			final DxfCellStyle poiDxfCellStyle = (DxfCellStyle) poiCellStyle;
			final Font poiFont = poiDxfCellStyle.getFont();
			return poiFont != null ? createDxfZssFont(poiFont) : null;
		} else {
			return super.importFont(poiCellStyle);
		}
	}
	
	//ZSS-1130
	@Override
	protected void importConditionalFormatting(SSheet sheet, Sheet poiSheet) {
		final XSSFSheet xssfSheet = (XSSFSheet) poiSheet;
		final List<XSSFConditionalFormatting> cfms = xssfSheet.getConditionalFormattings();
		if (cfms == null) return;
		
		for (XSSFConditionalFormatting cf : cfms) {
			SConditionalFormatting scf = prepareConditionalFormattingImpl(sheet, cf);
			((AbstractSheetAdv)sheet).addConditionalFormatting(scf);
		}
	}
	
	//ZSS-1130
	protected ConditionalFormattingImpl prepareConditionalFormattingImpl(SSheet sheet, XSSFConditionalFormatting cf) {
		final ConditionalFormattingImpl cfi = new ConditionalFormattingImpl(sheet);
		for (CellRangeAddress addr : cf.getFormattingRanges()) {
			final int r1 = addr.getFirstRow();
			final int c1 = addr.getFirstColumn();
			final int r2 = addr.getLastRow();
			final int c2 = addr.getLastColumn();
			cfi.addRegion(new CellRegion(r1, c1, r2, c2));
		}
		
		for (int j = 0, len = cf.getNumberOfRules(); j < len; ++j) {
			final XSSFConditionalFormattingRule rule = cf.getRule(j);
			cfi.addRule(prepareConditonalFormattingRuleImpl(rule));
		}
		return cfi;
	}
	
	//ZSS-1130
	protected ConditionalFormattingRuleImpl prepareConditonalFormattingRuleImpl(XSSFConditionalFormattingRule poiRule) {
		final ConditionalFormattingRuleImpl cfri = new ConditionalFormattingRuleImpl();
		CTCfRule ctRule = poiRule.getCTCfRule();
		cfri.setPriority(ctRule.getPriority());
		cfri.setType(toConditionalFormattingRuleType(ctRule.getType()));
		if (ctRule.isSetStopIfTrue())
			cfri.setStopIfTrue(ctRule.getStopIfTrue());
		if (ctRule.isSetDxfId())
			cfri.setExtraStyle(book.getExtraStyleAt((int)ctRule.getDxfId()));
		switch(cfri.getType()) {
		case ABOVE_AVERAGE:
			if (ctRule.isSetAboveAverage())
				cfri.setAboveAverage(ctRule.getAboveAverage());
			if (ctRule.isSetEqualAverage())
				cfri.setEqualAverage(ctRule.getEqualAverage());
			if (ctRule.isSetStdDev())
				cfri.setStandardDeviation(ctRule.getStdDev());
			break;
		case CELL_IS:
			if (ctRule.isSetOperator())
				cfri.setOperator(toCFRuleOperator(ctRule.getOperator()));
			prepareFormulas(cfri, ctRule);
			break;
		case COLOR_SCALE:
			cfri.setColorScale(prepareColorScale(ctRule));
			break;
		case CONTAINS_BLANKS:
		case NOT_CONTAINS_BLANKS:
			prepareFormulas(cfri, ctRule);
			break;
		case CONTAINS_ERRORS:
		case NOT_CONTAINS_ERRORS:
			prepareFormulas(cfri, ctRule);
			break;
		case CONTAINS_TEXT:
		case NOT_CONTAINS_TEXT:
			cfri.setText(ctRule.getText());
			cfri.setOperator(toCFRuleOperator(ctRule.getOperator()));
			prepareFormulas(cfri, ctRule);
			break;
		case DATA_BAR:
			cfri.setDataBar(prepareDataBar(ctRule));
			break;
		case BEGINS_WITH:
		case ENDS_WITH:
			cfri.setText(ctRule.getText());
			cfri.setOperator(toCFRuleOperator(ctRule.getOperator()));
			prepareFormulas(cfri, ctRule);
			break;
		case EXPRESSION:
			prepareFormulas(cfri, ctRule);
			break;
		case ICON_SET:
			cfri.setIconSet(prepareIconSet(ctRule));
			break;
		case TIME_PERIOD:
			cfri.setTimePeriod(toTimePeriod(ctRule.getTimePeriod()));
			prepareFormulas(cfri, ctRule);
			break;
		case TOP_10:
			if (ctRule.isSetRank())
				cfri.setRank(ctRule.getRank());
			if (ctRule.isSetPercent())
				cfri.setPercent(ctRule.getPercent());
			if (ctRule.isSetBottom())
				cfri.setBottom(ctRule.getBottom());
			break;
		case DUPLICATE_VALUES:
		case UNIQUE_VALUES:
			// No extra setting
			break;
		}
		
		return cfri;
	}
	
	//ZSS-1130
	protected SIconSet prepareIconSet(CTCfRule ctRule) {
		final IconSetImpl iconSet = new IconSetImpl();
		final CTIconSet ctIconSet = ctRule.getIconSet();
		for (CTCfvo ctvo : ctIconSet.getCfvoList()) {
			iconSet.addValueObject(prepareValueObject(ctvo));
		}
		iconSet.setType(toIconSetType(ctIconSet.getIconSet()));
		if (ctIconSet.isSetPercent())
			iconSet.setPercent(ctIconSet.getPercent());
		if (ctIconSet.isSetReverse())
			iconSet.setReverse(ctIconSet.getReverse());
		if (ctIconSet.isSetShowValue())
			iconSet.setShowValue(ctIconSet.getShowValue());
		
		return iconSet;
	}
	
	//ZSS-1130
	protected SColorScale prepareColorScale(CTCfRule ctRule) {
		final ColorScaleImpl colorScale = new ColorScaleImpl();
		final CTColorScale ctColorScale = ctRule.getColorScale();
		for (CTCfvo ctvo : ctColorScale.getCfvoList()) {
			colorScale.addValueObject(prepareValueObject(ctvo));
		}
		for (CTColor ctColor : ctColorScale.getColorList()) {
			final byte[] rgb = ctColor.getRgb();
			final SColor color = new ColorImpl(rgb); 
			colorScale.addColor(color);
		}
		return colorScale; 
	}
	
	//ZSS-1130
	protected SDataBar prepareDataBar(CTCfRule ctRule) {
		final DataBarImpl dataBar = new DataBarImpl();
		final CTDataBar ctDataBar = ctRule.getDataBar();
		for (CTCfvo ctvo : ctDataBar.getCfvoList()) {
			dataBar.addValueObject(prepareValueObject(ctvo));
		}
		final byte[] rgb = ctDataBar.getColor().getRgb();
		final SColor color = new ColorImpl(rgb); 
		dataBar.setColor(color);
		if (ctDataBar.isSetMaxLength())
			dataBar.setMaxLength(ctDataBar.getMaxLength());
		if (ctDataBar.isSetMinLength())
			dataBar.setMinLength(ctDataBar.getMinLength());
		if (ctDataBar.isSetShowValue())
			dataBar.setShowValue(ctDataBar.getShowValue());
		
		return dataBar;
	}

	//ZSS-1130
	protected void prepareFormulas(ConditionalFormattingRuleImpl cfri, CTCfRule ctRule) {
		for (String formula : ctRule.getFormulaList()) {
			cfri.addFormula(formula);
		}
	}
	
	//ZSS-1130
	protected SCFValueObject prepareValueObject(CTCfvo ctvo) {
		CFValueObjectImpl vo = new CFValueObjectImpl();
		if (ctvo.isSetGte())
			vo.setGreaterOrEqual(ctvo.getGte());
		vo.setType(toValueObjectType(ctvo.getType()));
		if (ctvo.isSetVal())
			vo.setValue(ctvo.getVal());
		
		return vo;
	}

	//ZSS-1130
	protected SConditionalFormattingRule.RuleTimePeriod toTimePeriod(STTimePeriod.Enum ctPeriod) {
		return SConditionalFormattingRule.RuleTimePeriod.values()[ctPeriod.intValue() - 1];
	}
	//ZSS-1130
	protected SConditionalFormattingRule.RuleOperator toCFRuleOperator(STConditionalFormattingOperator.Enum ctType) {
		return SConditionalFormattingRule.RuleOperator.values()[ctType.intValue()-1];
	}
	
	//ZSS-1130
	protected SIconSet.IconSetType toIconSetType(STIconSetType.Enum ctType) {
		return SIconSet.IconSetType.values()[ctType.intValue()-1];
	}
	
	//ZSS-1130
	protected SCFValueObject.CFValueObjectType toValueObjectType(STCfvoType.Enum ctType) {
		return SCFValueObject.CFValueObjectType.values()[ctType.intValue()-1];
	}
	
	//ZSS-1130
	protected SConditionalFormattingRule.RuleType toConditionalFormattingRuleType(STCfType.Enum cfType) {
		return RuleType.values()[cfType.intValue()-1];
	}
	
	//ZSS-1130
	protected STCfType.Enum toConditionalFormatingRuleType(SConditionalFormattingRule.RuleType stype) {
		return STCfType.Enum.forInt(stype.value);
	}

	//ZSS-992
	protected void importTableStyles() {
		((AbstractBookAdv)book).clearTableStyles();
		for (TableStyle poiStyle : workbook.getTableStyles()) {
			book.addTableStyle(importTableStyle(poiStyle));
		}
		book.setDefaultPivotStyleName(workbook.getDefaultPivotStyle());
		book.setDefaultTableStyleName(workbook.getDefaultTableStyle());
	}

	//ZSS-992
	protected STableStyleElem importTableStyleElem(TableStyle poiTableStyle, String name) {
		final DxfCellStyle dxfStyle = (DxfCellStyle) poiTableStyle.getDxfCellStyle(name);
		if (dxfStyle == null) return null;
		SExtraStyle extraStyle = book.getExtraStyleAt(dxfStyle.getIndex());
		return new TableStyleElemImpl(extraStyle.getFont(), extraStyle.getFill(), extraStyle.getBorder());
	}
	
	//ZSS-992
	protected STableStyle importTableStyle(TableStyle poiTableStyle) {
		final String name = poiTableStyle.getName();
		final STableStyleElem wholeTable = importTableStyleElem(poiTableStyle, "wholeTable");
		final STableStyleElem colStripe1 = importTableStyleElem(poiTableStyle, "firstColumnStripe");
		final STableStyleElem colStripe2 = importTableStyleElem(poiTableStyle, "SecondColumnStripe");
		final STableStyleElem rowStripe1 = importTableStyleElem(poiTableStyle, "firstRowStripe");
		final STableStyleElem rowStripe2 = importTableStyleElem(poiTableStyle, "SecondRowStripe");
		final STableStyleElem lastCol = importTableStyleElem(poiTableStyle, "lastColumn");
		final STableStyleElem firstCol = importTableStyleElem(poiTableStyle, "firstColumn");
		final STableStyleElem headerRow = importTableStyleElem(poiTableStyle, "headerRow");
		final STableStyleElem totalRow = importTableStyleElem(poiTableStyle, "totalRow");
		final STableStyleElem firstHeaderCell = importTableStyleElem(poiTableStyle, "firstHeaderCell");
		final STableStyleElem lastHeaderCell = importTableStyleElem(poiTableStyle, "lastHeaderCell");
		final STableStyleElem firstTotalCell = importTableStyleElem(poiTableStyle, "firstTotalCell");
		final STableStyleElem lastTotalCell = importTableStyleElem(poiTableStyle, "lastTotalCell");

		return new TableStyleImpl(
				name,
				wholeTable,
				colStripe1,
				1, 		//colStripe1Size,
				colStripe2,
				1, 		//colStripe2Size,
				rowStripe1,
				1, 		//rowStripe1Size,
				rowStripe2,
				1, 		//rowStripe2Size,
				lastCol,
				firstCol,
				headerRow,
				totalRow,
				firstHeaderCell,
				lastHeaderCell,
				firstTotalCell,
				lastTotalCell
		);
	}
	
	//ZSS-1191
	//@since 3.9.0
	@Override
	protected SColorFilter importColorFilter(ColorFilter colorFilter) {
		if (colorFilter == null) return null;
		final SExtraStyle src = importExtraStyle(colorFilter.getDxfCellStyle());
		final CellStyleMatcher matcher = new CellStyleMatcher(src);
		SExtraStyle style = book.searchExtraStyle(matcher);
		if (style == null) {
			book.addExtraStyle(src);
			style = src;
		}
		return new ColorFilterImpl(style, colorFilter.isByFontColor());
	}
}
 
