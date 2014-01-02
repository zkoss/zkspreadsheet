package org.zkoss.zss.ngmodel;

import static org.junit.Assert.*;

import java.net.URL;
import java.text.DateFormat;
import java.util.Locale;

import org.junit.Test;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NChart.*;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.chart.NGeneralChartData;

/**
 * Common test cases for importer & exporter.
 * @author kuro, hawk
 *
 */
public class ImExpTestBase {
	
	protected static final URL DEFAULT_CHART_IMPORT_FILE = ImporterTest.class.getResource("book/chart.xlsx");
	protected static String DEFAULT_BOOK_NAME = "PoiBook";
	protected static URL fileForImporterTest = ImporterTest.class.getResource("book/import2007.xlsx");


	protected void sheetTest(NBook book) {
		assertEquals(7, book.getNumOfSheet());

		NSheet sheet1 = book.getSheetByName("Value");
		assertEquals("Value", sheet1.getSheetName());
		assertEquals(20, sheet1.getDefaultRowHeight());
		assertEquals(64, sheet1.getDefaultColumnWidth());
		
		NSheet sheet2 = book.getSheetByName("Style");
		assertEquals("Style", sheet2.getSheetName());
		NSheet sheet3 = book.getSheetByName("NamedRange");
		assertEquals("NamedRange", sheet3.getSheetName());
	}	
	
	
	protected void sheetProtectionTest(NBook book) {
		assertFalse(book.getSheetByName("Value").isProtected());
		assertTrue(book.getSheetByName("sheet-protection").isProtected());
	}
	
	
	protected void sheetNamedRangeTest(NBook book) {
		assertEquals(2, book.getNumOfName());
		assertEquals("NamedRange!$B$2:$D$3", book.getNameByName("TestRange1", "NamedRange").getRefersToFormula());
		assertEquals("NamedRange!$F$2", book.getNameByName("RangeMerged", "NamedRange").getRefersToFormula());
	}

	
	protected void cellValueTest(NBook book) {
		NSheet sheet = book.getSheetByName("Value");
		//text
		assertEquals(NCell.CellType.STRING, sheet.getCell(0,1).getType());
		assertEquals("B1", sheet.getCell(0,1).getStringValue());
		assertEquals("C1", sheet.getCell(0,2).getStringValue());
		assertEquals("D1", sheet.getCell(0,3).getStringValue());
		
		//number
		assertEquals(NCell.CellType.NUMBER, sheet.getCell(1,1).getType());
		assertEquals(123, sheet.getCell(1,1).getNumberValue().intValue());
		assertEquals(123.45, sheet.getCell(1,2).getNumberValue().doubleValue(), 0.01);
		
		//date
		assertEquals(NCell.CellType.NUMBER, sheet.getCell(2,1).getType());
		assertEquals(41618, sheet.getCell(2,1).getNumberValue().intValue());
		assertEquals("Dec 10, 2013", DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,1).getDateValue()));
		assertEquals(0.61, sheet.getCell(2,2).getNumberValue().doubleValue(), 0.01);
		assertEquals("2:44:10 PM", DateFormat.getTimeInstance (DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,2).getDateValue()));
		
		//formula
		assertEquals(NCell.CellType.FORMULA, sheet.getCell(3,1).getType());
		assertEquals("SUM(10,20)", sheet.getCell(3,1).getFormulaValue());
		assertEquals("ISBLANK(B1)", sheet.getCell(3,2).getFormulaValue());
		assertEquals("B1", sheet.getCell(3,3).getFormulaValue());
		
		//error
		assertEquals(NCell.CellType.ERROR, sheet.getCell(4,1).getType());
		assertEquals(ErrorValue.INVALID_NAME, sheet.getCell(4,1).getErrorValue().getCode());
		assertEquals(ErrorValue.INVALID_VALUE, sheet.getCell(4,2).getErrorValue().getCode());
		
		//blank
		assertEquals(NCell.CellType.BLANK, sheet.getCell(5,1).getType());
		assertEquals("", sheet.getCell(5,1).getStringValue());
	}
	
	
	protected void cellStyleTest(NBook book){
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(true, sheet.getCell(24, 1).getCellStyle().isWrapText());
		//alignment
		assertEquals(VerticalAlignment.TOP, sheet.getCell(26, 1).getCellStyle().getVerticalAlignment());
		assertEquals(VerticalAlignment.CENTER, sheet.getCell(26, 2).getCellStyle().getVerticalAlignment());
		assertEquals(VerticalAlignment.BOTTOM, sheet.getCell(26, 3).getCellStyle().getVerticalAlignment());
		assertEquals(Alignment.LEFT, sheet.getCell(27, 1).getCellStyle().getAlignment());
		assertEquals(Alignment.CENTER, sheet.getCell(27, 2).getCellStyle().getAlignment());
		assertEquals(Alignment.RIGHT, sheet.getCell(27, 3).getCellStyle().getAlignment());
		//cell filled color
		assertEquals("#FF0000", sheet.getCell(11, 0).getCellStyle().getFillColor().getHtmlColor());
		assertEquals("#00FF00", sheet.getCell(11, 1).getCellStyle().getFillColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(11, 2).getCellStyle().getFillColor().getHtmlColor());

		//ensure cell style reusing
		assertTrue(sheet.getCell(27, 0).getCellStyle().equals(sheet.getCell(26, 0).getCellStyle()));
		assertTrue(sheet.getCell(28, 0).getCellStyle().equals(sheet.getCell(26, 0).getCellStyle()));
		assertTrue(sheet.getCell(28, 0).getCellStyle().getFont().equals(sheet.getCell(26, 0).getCellStyle().getFont()));
		
		//fill pattern
		assertEquals(FillPattern.SOLID_FOREGROUND, sheet.getCell(37, 1).getCellStyle().getFillPattern());
		assertEquals(FillPattern.ALT_BARS, sheet.getCell(37, 2).getCellStyle().getFillPattern());
		
		NSheet protectedSheet = book.getSheetByName("sheet-protection");
		assertEquals(true, protectedSheet.getCell(0, 0).getCellStyle().isLocked());
		assertEquals(false, protectedSheet.getCell(1, 0).getCellStyle().isLocked());
	}

	
	protected void cellBorderTest(NBook book){
		NSheet sheet = book.getSheetByName("cell-border");
		assertEquals(BorderType.NONE, sheet.getCell(2, 0).getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, sheet.getCell(2, 1).getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, sheet.getCell(2, 2).getCellStyle().getBorderTop());
		assertEquals(BorderType.THIN, sheet.getCell(2, 3).getCellStyle().getBorderLeft());
		assertEquals(BorderType.THIN, sheet.getCell(2, 4).getCellStyle().getBorderRight());
		
		assertEquals(BorderType.THIN, sheet.getCell(4, 1).getCellStyle().getBorderBottom());
		assertEquals(BorderType.DOTTED, sheet.getCell(4, 2).getCellStyle().getBorderBottom());
		assertEquals(BorderType.DASHED, sheet.getCell(4, 3).getCellStyle().getBorderBottom());
		
		assertEquals("#FF0000", sheet.getCell(14, 1).getCellStyle().getBorderBottomColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(14, 1).getCellStyle().getBorderLeftColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(14, 1).getCellStyle().getBorderTopColor().getHtmlColor());
		assertEquals("#FF0000", sheet.getCell(14, 1).getCellStyle().getBorderRightColor().getHtmlColor());
	}

	
	protected void cellFontNameTest(NBook book){
		NSheet sheet = book.getSheetByName("Style");
		assertEquals("Arial", sheet.getCell(3, 0).getCellStyle().getFont().getName());
		assertEquals("Arial Black", sheet.getCell(3, 1).getCellStyle().getFont().getName());
		assertEquals("Calibri", sheet.getCell(3, 2).getCellStyle().getFont().getName());
	}
	
	
	protected void cellFontStyleTest(NBook book){
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(NFont.Boldweight.BOLD, sheet.getCell(9, 0).getCellStyle().getFont().getBoldweight());
		assertTrue(sheet.getCell(9, 1).getCellStyle().getFont().isItalic());
		assertTrue(sheet.getCell(9, 2).getCellStyle().getFont().isStrikeout());
		assertEquals(NFont.Underline.SINGLE, sheet.getCell(9, 3).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.DOUBLE, sheet.getCell(9, 4).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.SINGLE_ACCOUNTING, sheet.getCell(9, 5).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.DOUBLE_ACCOUNTING, sheet.getCell(9, 6).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.NONE, sheet.getCell(9, 7).getCellStyle().getFont().getUnderline());
		
		//height
		assertEquals(8, sheet.getCell(6, 0).getCellStyle().getFont().getHeightPoints());
		assertEquals(72, sheet.getCell(6, 3).getCellStyle().getFont().getHeightPoints());
		
		//type offset
		assertEquals(TypeOffset.SUPER, sheet.getCell(32, 1).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.SUB, sheet.getCell(32, 2).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.NONE, sheet.getCell(32, 3).getCellStyle().getFont().getTypeOffset());
	}
	
	
	protected void cellFontColorTest(NBook book){
		NSheet sheet = book.getSheetByName("Style");
		assertEquals("#000000", sheet.getCell(0, 0).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#FF0000", sheet.getCell(1, 0).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#00FF00", sheet.getCell(1, 1).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#0000FF", sheet.getCell(1, 2).getCellStyle().getFont().getColor().getHtmlColor());
	}
	
	
	protected void rowTest(NBook book){
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(28, sheet.getRow(0).getHeight());
		assertEquals(20, sheet.getRow(1).getHeight());
		//style
		NCellStyle rowStyle1 = sheet.getRow(34).getCellStyle();
		assertEquals("#0000FF",rowStyle1.getFont().getColor().getHtmlColor());
		assertEquals(12,rowStyle1.getFont().getHeightPoints());
		NCellStyle rowStyle2 = sheet.getRow(35).getCellStyle();
		assertEquals(true,rowStyle2.getFont().isItalic());
		assertEquals(14,rowStyle2.getFont().getHeightPoints());		
		
		NSheet rowSheet = book.getSheetByName("column-row");
		assertTrue(rowSheet.getRow(9).isHidden());
		assertTrue(rowSheet.getRow(10).isHidden());
	}

	/**
	 * Information technology ??Document description and processing languages ??
	 * Office Open XML File Formats ??Part 1: Fundamentals and Markup LanguageReference  
	 * 18.8.30 numFmt (Number Format) 
	 */
	
	protected void cellFormatTest(NBook book){
		NSheet sheet = book.getSheetByName("Format");
		assertEquals("#,##0.00", sheet.getCell(1, 1).getCellStyle().getDataFormat());
		assertEquals("\"NT$\"#,##0.00", sheet.getCell(1, 2).getCellStyle().getDataFormat());
		assertEquals("m/d/yyyy", sheet.getCell(1, 4).getCellStyle().getDataFormat());
		//Excel shows "hh:mm AM/PM"
		assertEquals("h:mm AM/PM", sheet.getCell(1, 5).getCellStyle().getDataFormat());
		assertEquals("0.0%", sheet.getCell(1, 6).getCellStyle().getDataFormat());
		assertEquals("# ??/??", sheet.getCell(3, 1).getCellStyle().getDataFormat());
		assertEquals("0.00E+00", sheet.getCell(3, 2).getCellStyle().getDataFormat());
		assertEquals("@", sheet.getCell(3, 3).getCellStyle().getDataFormat());
		assertEquals("[<=9999999]###\\-####;\\(0#\\)\\ ###\\-####", sheet.getCell(3, 4).getCellStyle().getDataFormat());
	}


	protected void columnTest(NBook book){
		NSheet sheet = book.getSheetByName("column-row");
		//column style
		assertEquals(NFont.Boldweight.BOLD, sheet.getColumn(0).getCellStyle().getFont().getBoldweight());
		//width
		assertEquals(228, sheet.getColumn(0).getWidth()); 
		assertEquals(100, sheet.getColumn(1).getWidth());
		assertEquals(102, sheet.getColumn(2).getWidth());
		assertEquals(64, sheet.getColumn(4).getWidth());	//the hidden column
		assertEquals(64, sheet.getColumn(5).getWidth());	//default width
		
		//the hidden column
		assertFalse(sheet.getColumn(3).isHidden());
		assertTrue(sheet.getColumn(4).isHidden());		
	}
	
	/**
	 * import last column that only has column width change but has all empty cells 
	 */
	
	protected void lastChangedColumnTest(NBook book){
		NSheet sheet = book.getSheetByName("column-row");
		assertEquals(110, sheet.getColumn(13).getWidth());
	}
	
	
	protected void viewInfoTest(NBook book){
		NSheet sheet = book.getSheetByName("column-row");
		assertEquals(3, sheet.getViewInfo().getNumOfRowFreeze());
		assertEquals(1, sheet.getViewInfo().getNumOfColumnFreeze());
		
		//grid line display
		assertTrue(sheet.getViewInfo().isDisplayGridline());
		assertFalse(book.getSheetByName("cell-border").getViewInfo().isDisplayGridline());
	}

	
	protected void mergedTest(NBook book){
		NSheet sheet = book.getSheetByName("cell-border");
		assertEquals(4, sheet.getMergedRegions().size());
		assertEquals("C31:D33", sheet.getMergedRegions().get(0).getReferenceString());
		assertEquals("B31:B33", sheet.getMergedRegions().get(1).getReferenceString());
		assertEquals("E28:G28", sheet.getMergedRegions().get(2).getReferenceString());
		assertEquals("B28:C28", sheet.getMergedRegions().get(3).getReferenceString());
	}	
	
	protected void areaChart(NBook book){
		NSheet sheet = book.getSheetByName("Area");
		NChart areaChart = sheet.getChart(0);
		assertEquals(NChartType.AREA,areaChart.getType());
		
		//a chart locating in one column and one row test
		assertEquals(493, areaChart.getAnchor().getWidth());
		assertEquals(283, areaChart.getAnchor().getHeight());
		
		NGeneralChartData chartData = (NGeneralChartData)areaChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
		
		NChart area3dChart = sheet.getChart(1);
		assertEquals(NChartGrouping.STANDARD, area3dChart.getGrouping());
		assertEquals(NChartLegendPosition.BOTTOM, area3dChart.getLegendPosition());
	}
	
	protected void barChart(NBook book){
		NSheet sheet = book.getSheetByName("Bar");
		NChart barChart = sheet.getChart(0);
		
		assertEquals(NChartType.BAR,barChart.getType());
		
		assertEquals(480, barChart.getAnchor().getWidth());
		assertEquals(284, barChart.getAnchor().getHeight());
		assertEquals(25, barChart.getAnchor().getXOffset());
		assertEquals(7, barChart.getAnchor().getYOffset());
		
		assertEquals(NBarDirection.HORIZONTAL, barChart.getBarDirection());
		assertEquals(NChartGrouping.CLUSTERED, barChart.getGrouping());
		assertEquals(false, barChart.isThreeD());
		assertEquals(NChartLegendPosition.RIGHT, barChart.getLegendPosition());
		
		//data
		NGeneralChartData chartData = (NGeneralChartData)barChart.getData();
		assertEquals(3, chartData.getNumOfCategory());
		assertEquals("Internet Explorer", chartData.getCategory(0));
		assertEquals("Chrome", chartData.getCategory(1));
		assertEquals("Firefox", chartData.getCategory(2));
		assertEquals(3, chartData.getNumOfSeries());
		assertEquals("January 2012", chartData.getSeries(0).getName());
		assertEquals(0.3427, chartData.getSeries(0).getValue(0));
		assertEquals(0.2599, chartData.getSeries(0).getValue(1));
		assertEquals(0.2268, chartData.getSeries(0).getValue(2));
		assertEquals("February 2012", chartData.getSeries(1).getName());
		assertEquals(0.327, chartData.getSeries(1).getValue(0));
		assertEquals(0.2724, chartData.getSeries(1).getValue(1));
		assertEquals(0.2276, chartData.getSeries(1).getValue(2));
		assertEquals("March 2012", chartData.getSeries(2).getName());
		assertEquals(0.3168, chartData.getSeries(2).getValue(0));
		assertEquals(0.2809, chartData.getSeries(2).getValue(1));
		assertEquals(0.2273, chartData.getSeries(2).getValue(2));
		
		NChart barChart3D = sheet.getChart(1);
		assertEquals(true, barChart3D.isThreeD());
	}
	
	public void bubbleChart(NBook book){
		NSheet sheet = book.getSheetByName("Bubble");
		NChart bubbleChart = sheet.getChart(0);
		assertEquals(NChartType.BUBBLE, bubbleChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)bubbleChart.getData();
		assertEquals(0, chartData.getNumOfCategory());
		assertEquals(1, chartData.getNumOfSeries());
		assertEquals("String Literal Title", chartData.getSeries(0).getName());
	}
	
	public void columnChart(NBook book){
		NSheet sheet = book.getSheetByName("Column");
		NChart columnChart = sheet.getChart(0);
		assertEquals(NChartType.COLUMN,columnChart.getType());
		assertEquals(NBarDirection.VERTICAL, columnChart.getBarDirection());
		assertEquals(NChartLegendPosition.TOP, columnChart.getLegendPosition());
		
		NGeneralChartData chartData = (NGeneralChartData)columnChart.getData();
		assertEquals(4, chartData.getNumOfCategory());
		
		NChart column3dChart = sheet.getChart(1);
		assertEquals(NChartGrouping.STACKED, column3dChart.getGrouping());
	}
	
	public void doughnutChart(NBook book){
		NSheet sheet = book.getSheetByName("Doughnut");
		NChart doughnutChart = sheet.getChart(0);
		assertEquals(NChartType.DOUGHNUT, doughnutChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)doughnutChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
	}
	
	public void lineChart(NBook book){
		NSheet sheet = book.getSheetByName("Line");
		NChart lineChart = sheet.getChart(0);
		assertEquals(NChartType.LINE, lineChart.getType());
		NGeneralChartData chartData = (NGeneralChartData)lineChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
		
		NChart line3dChart = sheet.getChart(1);
		assertEquals(true, line3dChart.isThreeD());
		chartData = (NGeneralChartData)line3dChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
	}
	
	public void pieChart(NBook book){
		NSheet sheet = book.getSheetByName("Pie");
		NChart pieChart = sheet.getChart(0);
		assertEquals(NChartType.PIE, pieChart.getType());
		assertEquals(null,pieChart.getTitle());
		NGeneralChartData chartData = (NGeneralChartData)pieChart.getData();
		assertEquals(1, chartData.getNumOfSeries());
		
		NChart pie3dChart = sheet.getChart(1);
		assertEquals(NChartType.PIE, pie3dChart.getType());
		assertEquals(true, pie3dChart.isThreeD());
	}
	
	public void scatterChart(NBook book){
		NSheet sheet = book.getSheetByName("Scatter");
		NChart scatterChart = sheet.getChart(0);
		assertEquals(NChartType.SCATTER, scatterChart.getType());
		
		NGeneralChartData chartData = (NGeneralChartData)scatterChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
		assertEquals("Internet Explorer", chartData.getSeries(0).getName());
		assertEquals(0.3427, chartData.getSeries(0).getYValue(0));
		assertEquals(0.327, chartData.getSeries(0).getYValue(1));
		assertEquals(0.3168, chartData.getSeries(0).getYValue(2));
	}
	
}

