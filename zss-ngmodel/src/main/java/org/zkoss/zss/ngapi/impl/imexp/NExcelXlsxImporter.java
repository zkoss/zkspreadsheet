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
import java.util.logging.*;

import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.w3c.dom.Node;
import org.zkoss.poi.POIXMLDocumentPart;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.usermodel.charts.*;
import org.zkoss.poi.xssf.model.ExternalLink;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.poi.xssf.usermodel.charts.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NChart.NBarDirection;
import org.zkoss.zss.ngmodel.NChart.NChartGrouping;
import org.zkoss.zss.ngmodel.NChart.NChartLegendPosition;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.NPicture.Format;
import org.zkoss.zss.ngmodel.chart.*;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
/**
 * Specific importing behavior for XLSX.
 * 
 * @author Hawk
 * @since 3.5.0
 */
public class NExcelXlsxImporter extends AbstractExcelImporter{
	private static final Logger logger = Logger.getLogger(NExcelXlsxImporter.class.getName());

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
	protected void importColumn(Sheet poiSheet, NSheet sheet, int defaultWidth) {
		CTWorksheet worksheet = ((XSSFSheet)poiSheet).getCTWorksheet();
		if(worksheet.sizeOfColsArray()<=0){
			return;
		}
		
		CTCols colsArray = worksheet.getColsArray(0);
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
			CTCol ctCol = colsArray.getColArray(i);
			//max is 16384
			
			NColumnArray columnArray = sheet.setupColumnArray((int)ctCol.getMin()-1, (int)ctCol.getMax()-1);
			boolean hidden = ctCol.getHidden();
			int columnIndex = (int)ctCol.getMin()-1;
			
			columnArray.setHidden(hidden);
			if (hidden == false){
				//when CT_Col is hidden with default width, We don't import the width for it's 0.  
				columnArray.setWidth(XUtils.getWidthAny(poiSheet, columnIndex, CHRACTER_WIDTH));
			}
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
	private void importChart(List<ZssChartX> poiCharts, Sheet poiSheet, NSheet sheet) {
		
		for (ZssChartX zssChart : poiCharts){
			XSSFChart xssfChart = (XSSFChart)zssChart.getChart();
			NViewAnchor viewAnchor = toViewAnchor(poiSheet, xssfChart.getPreferredSize());

			NChart chart = null;
			CategoryData categoryData = null;
			//reference ChartHelper.drawXSSFChart()
			switch(xssfChart.getChartType()){
				case Area:
					chart = sheet.addChart(NChartType.AREA, viewAnchor);
					categoryData = new XSSFAreaChartData(xssfChart);
					chart.setGrouping(convertGrouping(((XSSFAreaChartData)categoryData).getGrouping()));
					break;
				case Area3D:
					chart = sheet.addChart(NChartType.AREA, viewAnchor);
					categoryData = new XSSFArea3DChartData(xssfChart);
					chart.setGrouping(convertGrouping(((XSSFArea3DChartData)categoryData).getGrouping()));
					break;
				case Bar:
					chart = sheet.addChart(NChartType.BAR, viewAnchor);
					categoryData = new XSSFBarChartData(xssfChart);
					chart.setBarDirection(convertBarDirection(((XSSFBarChartData)categoryData).getBarDirection()));
					chart.setGrouping(convertGrouping(((XSSFBarChartData)categoryData).getGrouping()));
					break;
				case Bar3D:
					chart = sheet.addChart(NChartType.BAR, viewAnchor);
					categoryData = new XSSFBar3DChartData(xssfChart);
					chart.setBarDirection(convertBarDirection(((XSSFBar3DChartData)categoryData).getBarDirection()));
					chart.setGrouping(convertGrouping(((XSSFBar3DChartData)categoryData).getGrouping()));
					break;
				case Bubble:
					chart = sheet.addChart(NChartType.BUBBLE, viewAnchor);
					
					XYZData xyzData = new XSSFBubbleChartData(xssfChart);
					importXyzSeries(xyzData.getSeries(), (NGeneralChartData)chart.getData());
					break;
				case Column:
					chart = sheet.addChart(NChartType.COLUMN, viewAnchor);
					categoryData = new XSSFColumnChartData(xssfChart);
					chart.setBarDirection(convertBarDirection(((XSSFColumnChartData)categoryData).getBarDirection()));
					chart.setGrouping(convertGrouping(((XSSFColumnChartData)categoryData).getGrouping()));
					break;
				case Column3D:
					chart = sheet.addChart(NChartType.COLUMN, viewAnchor);
					categoryData = new XSSFColumn3DChartData(xssfChart);
					chart.setBarDirection(convertBarDirection(((XSSFColumn3DChartData)categoryData).getBarDirection()));
					chart.setGrouping(convertGrouping(((XSSFColumn3DChartData)categoryData).getGrouping()));
					break;
				case Doughnut:
					chart = sheet.addChart(NChartType.DOUGHNUT, viewAnchor);
					categoryData = new XSSFDoughnutChartData(xssfChart);
					break;
				case Line:
					chart = sheet.addChart(NChartType.LINE, viewAnchor);
					categoryData = new XSSFLineChartData(xssfChart);
					break;
				case Line3D:
					chart = sheet.addChart(NChartType.LINE, viewAnchor);
					categoryData = new XSSFLine3DChartData(xssfChart);
					break;
				case Pie:
					chart = sheet.addChart(NChartType.PIE, viewAnchor);
					categoryData = new XSSFPieChartData(xssfChart);
					break;
				case Pie3D:
					chart = sheet.addChart(NChartType.PIE, viewAnchor);
					categoryData = new XSSFPie3DChartData(xssfChart);
					break;
				case Scatter:
					chart = sheet.addChart(NChartType.SCATTER, viewAnchor);
					
					XYData xyData =  new XSSFScatChartData(xssfChart);
					importXySeries(xyData.getSeries(), (NGeneralChartData)chart.getData());
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
			chart.setLegendPosition(convertLengendPosition(xssfChart.getOrCreateLegend().getPosition()));
			if (categoryData != null){
				importSeries(categoryData.getSeries(), (NGeneralChartData)chart.getData());
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
	private void importSeries(List<? extends CategoryDataSerie> seriesList, NGeneralChartData chartData) {
		CategoryDataSerie firstSeries = null;
		if ((firstSeries = seriesList.get(0))!=null){
			chartData.setCategoriesFormula(getValueFormula(firstSeries.getCategories()));
		}
		for (int i =0 ;  i< seriesList.size() ; i++){
			CategoryDataSerie sourceSeries = seriesList.get(i);
			String nameExpression = getTitleFormula(sourceSeries.getTitle(), i);			
			String xValueExpression = getValueFormula(sourceSeries.getValues());
			NSeries series = chartData.addSeries();
			series.setFormula(nameExpression, xValueExpression);
		}
	}
	
	private void importXySeries(List<? extends XYDataSerie> seriesList, NGeneralChartData chartData) {
		for (int i =0 ;  i< seriesList.size() ; i++){
			XYDataSerie sourceSeries = seriesList.get(i);
			NSeries series = chartData.addSeries();
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
	private void importXyzSeries(List<? extends XYZDataSerie> seriesList, NGeneralChartData chartData) {
		for (int i =0 ;  i< seriesList.size() ; i++){
			XYZDataSerie sourceSeries = seriesList.get(i);
			//reference to ChartHelper.prepareTitle()
			String nameExpression = getTitleFormula(sourceSeries.getTitle(), i);			
			String xValueExpression = getValueFormula(sourceSeries.getXs());
			String yValueExpression = getValueFormula(sourceSeries.getYs());
			String zValueExpression = getValueFormula(sourceSeries.getZs());	
			
			NSeries series = chartData.addSeries();
			series.setXYZFormula(nameExpression, xValueExpression, yValueExpression, zValueExpression);
		}
	}
	
	/**
	 * return a formula or generate a default title ("Series[N]")if title doesn't exist.
	 * reference ChartHelper.prepareTitle()
	 * @param textSource
	 * @return
	 */
	private String getTitleFormula(ChartTextSource textSource, int seriesIndex){
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
	private String getValueFormula(ChartDataSource<?> dataSource){
		if (dataSource.isReference()){
			return dataSource.getFormulaString();
		}else{
			StringBuilder expression = new StringBuilder("{");
			int count = dataSource.getPointCount();
			for (int i = 0; i < count; ++i) {
				Object value = dataSource.getPointAt(i); //Number or String
				if (value == null){
					if (dataSource.isNumeric()){
						expression.append("0");
					}else{
						expression.append("\"\"");
					}
				}else{
					expression.append(value.toString());
				}
				if (i != count-1){
					expression.append(",");
				}
			}
			expression.append("}");
			return expression.toString();
		}
	}
	
	
	private NChartLegendPosition convertLengendPosition(LegendPosition position){
		switch(position){
		case BOTTOM:
			return NChartLegendPosition.BOTTOM;
		case LEFT:
			return NChartLegendPosition.LEFT;
		case TOP:
			return NChartLegendPosition.TOP;
		case TOP_RIGHT:
			return NChartLegendPosition.TOP_RIGHT;
		case RIGHT:
		default:
			return NChartLegendPosition.RIGHT;
		}
		
	}
	private NBarDirection convertBarDirection(ChartDirection direction){
		switch(direction){
		case VERTICAL:
			return NBarDirection.VERTICAL;
		case HORIZONTAL:
		default:
			return NBarDirection.HORIZONTAL;
		}
	}
	
	private NChartGrouping convertGrouping(ChartGrouping grouping){
		switch(grouping){
		case STACKED:
			return NChartGrouping.STACKED;
		case PERCENT_STACKED:
			return NChartGrouping.PERCENT_STACKED;
		case CLUSTERED:
			return NChartGrouping.CLUSTERED;
		case STANDARD:
		default:
			return NChartGrouping.STANDARD;
		}
	}
	
	/*
	 * import drawings which include charts and pictures.
	 * reference DrawingManagerImpl.initXSSFDrawings()
	 */
	@Override
	protected void importDrawings(Sheet poiSheet, NSheet sheet) {
		/**
		 * A list of POI chart wrapper loaded during import.
		 */
		List<ZssChartX> poiCharts = new LinkedList<ZssChartX>();
		List<Picture> poiPictures = new LinkedList<Picture>();
		
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
	
	private void importPicture(List<Picture> poiPictures, Sheet poiSheet, NSheet sheet){
		for (Picture picture : poiPictures){
			System.out.println(picture.getPictureData().getMimeType());
			sheet.addPicture(Format.PNG, picture.getPictureData().getData(), toViewAnchor(poiSheet, picture.getClientAnchor()));
		}
	}
	
	private XSSFChartX createXSSFChartX(XSSFDrawing patriarch, CTGraphicalObjectFrame gfrm , XSSFClientAnchor xanchor) {
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
		final int firstColumn = anchor.getCol1();
	    final int firstColumnWidth = XUtils.getWidthAny(poiSheet,firstColumn, AbstractExcelImporter.CHRACTER_WIDTH);
	    int offsetInFirstColumn = UnitUtil.emuToPx(anchor.getDx1());

	    final int anchorWidthInFirstColumn = firstColumnWidth - offsetInFirstColumn;  
		int anchorWidthInLastColumn = UnitUtil.emuToPx(anchor.getDx2());
	    
		final int lastColumn = anchor.getCol2();
	    //in between
	    int width = firstColumn == lastColumn ? anchorWidthInLastColumn - offsetInFirstColumn : anchorWidthInFirstColumn + anchorWidthInLastColumn;
	    width = Math.abs(width); // just in case
	    
	    for (int j = firstColumn+1; j < lastColumn; ++j) {
	    	width += XUtils.getWidthAny(poiSheet,j, AbstractExcelImporter.CHRACTER_WIDTH);
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
	    final int firstRowHeight = XUtils.getHeightAny(poiSheet,firstRow);
		final int anchorHeightInFirstRow = firstRowHeight - offsetInFirstRow;  
	    
	    //last row
		final int lastRow = anchor.getRow2();
		int anchorHeightInLastRow = UnitUtil.emuToPx(anchor.getDy2());
		int height = lastRow == firstRow ? anchorHeightInLastRow - offsetInFirstRow : anchorHeightInFirstRow + anchorHeightInLastRow ;
		height = Math.abs(height); // just in case
	    
	    //add inter-row height
	    for (int row = firstRow+1; row < lastRow; ++row) {
	    	height += XUtils.getHeightAny(poiSheet,row);
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
}
 