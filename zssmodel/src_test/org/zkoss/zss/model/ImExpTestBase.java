package org.zkoss.zss.model;

import static org.junit.Assert.*;

import java.net.URL;
import java.text.DateFormat;
import java.util.Locale;

import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SHyperlink;
import org.zkoss.zss.model.SPicture;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SCellStyle.Alignment;
import org.zkoss.zss.model.SCellStyle.BorderType;
import org.zkoss.zss.model.SCellStyle.FillPattern;
import org.zkoss.zss.model.SCellStyle.VerticalAlignment;
import org.zkoss.zss.model.SChart.*;
import org.zkoss.zss.model.SDataValidation.ErrorStyle;
import org.zkoss.zss.model.SDataValidation.OperatorType;
import org.zkoss.zss.model.SDataValidation.ValidationType;
import org.zkoss.zss.model.SFont.TypeOffset;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.model.chart.SGeneralChartData;

/**
 * Common test cases for importer & exporter.
 * @author kuro, hawk
 *
 */
public class ImExpTestBase {
	/**
	 * We create all XLSX file with Excel 2007. (With 2010, it will have differences.) 
	 */
	protected URL IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/import.xlsx");
	protected URL CHART_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/chart.xlsx");
	protected URL PICTURE_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/picture.xlsx");
	protected URL FILTER_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/filter.xlsx");
	protected static String DEFAULT_BOOK_NAME = "PoiBook";
	
	protected void hyperlinkTest(SBook book) {
		SSheet sheet = book.getSheetByName("Style");
		SCell cell = sheet.getCell("B31");
		SHyperlink link = cell.getHyperlink();
		
		assertEquals("http://www.zkoss.org/", link.getAddress());
		assertEquals("", link.getLabel());
		assertEquals(SHyperlink.HyperlinkType.URL, link.getType());
	}
	
	protected void sheetTest(SBook book) {
		assertEquals(8, book.getNumOfSheet());

		SSheet sheet1 = book.getSheetByName("Value");
		assertEquals("Value", sheet1.getSheetName());
		assertEquals(20, sheet1.getDefaultRowHeight());
		assertEquals(64, sheet1.getDefaultColumnWidth());
		
		SSheet sheet2 = book.getSheetByName("Style");
		assertEquals("Style", sheet2.getSheetName());
		SSheet sheet3 = book.getSheetByName("NamedRange");
		assertEquals("NamedRange", sheet3.getSheetName());
	}	
	
	
	protected void sheetProtectionTest(SBook book) {
		assertFalse(book.getSheetByName("Value").isProtected());
		assertTrue(book.getSheetByName("sheet-protection").isProtected());
	}
	
	
	protected void sheetNamedRangeTest(SBook book) {
		assertEquals(2, book.getNumOfName());
		assertEquals("NamedRange!$B$2:$D$3", book.getNameByName("TestRange1").getRefersToFormula());
		assertEquals("NamedRange!$F$2", book.getNameByName("RangeMerged").getRefersToFormula());
	}

	
	protected void cellValueTest(SBook book) {
		SSheet sheet = book.getSheetByName("Value");
		//text
		assertEquals(SCell.CellType.STRING, sheet.getCell(0,1).getType());
		assertEquals("B1", sheet.getCell(0,1).getStringValue());
		assertEquals("C1", sheet.getCell(0,2).getStringValue());
		assertEquals("D1", sheet.getCell(0,3).getStringValue());
		
		//number
		assertEquals(SCell.CellType.NUMBER, sheet.getCell(1,1).getType());
		assertEquals(123, sheet.getCell(1,1).getNumberValue().intValue());
		assertEquals(123.45, sheet.getCell(1,2).getNumberValue().doubleValue(), 0.01);
		
		//date
		assertEquals(SCell.CellType.NUMBER, sheet.getCell(2,1).getType());
		assertEquals(41618, sheet.getCell(2,1).getNumberValue().intValue());
		assertEquals("Dec 10, 2013", DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,1).getDateValue()));
		assertEquals(0.61, sheet.getCell(2,2).getNumberValue().doubleValue(), 0.01);
		assertEquals("2:44:10 PM", DateFormat.getTimeInstance (DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,2).getDateValue()));
		
		//formula
		assertEquals(SCell.CellType.FORMULA, sheet.getCell(3,1).getType());
		assertEquals("SUM(10,20)", sheet.getCell(3,1).getFormulaValue());
		assertEquals("ISBLANK(B1)", sheet.getCell(3,2).getFormulaValue());
		assertEquals("B1", sheet.getCell(3,3).getFormulaValue());
		
		//error
		assertEquals(SCell.CellType.ERROR, sheet.getCell(4,1).getType());
		assertEquals(ErrorValue.INVALID_NAME, sheet.getCell(4,1).getErrorValue().getCode());
		assertEquals(ErrorValue.INVALID_VALUE, sheet.getCell(4,2).getErrorValue().getCode());
		
		//blank
		assertEquals(SCell.CellType.BLANK, sheet.getCell(5,1).getType());
		assertEquals("", sheet.getCell(5,1).getStringValue());
	}
	
	
	protected void cellStyleTest(SBook book){
		SSheet sheet = book.getSheetByName("Style");
		assertEquals(true, sheet.getCell(24, 1).getCellStyle().isWrapText());
		//alignment
		assertEquals(VerticalAlignment.TOP, sheet.getCell(26, 1).getCellStyle().getVerticalAlignment());
		assertEquals(VerticalAlignment.CENTER, sheet.getCell(26, 2).getCellStyle().getVerticalAlignment());
		assertEquals(VerticalAlignment.BOTTOM, sheet.getCell(26, 3).getCellStyle().getVerticalAlignment());
		assertEquals(Alignment.LEFT, sheet.getCell(27, 1).getCellStyle().getAlignment());
		assertEquals(Alignment.CENTER, sheet.getCell(27, 2).getCellStyle().getAlignment());
		assertEquals(Alignment.RIGHT, sheet.getCell(27, 3).getCellStyle().getAlignment());
		//cell filled color
		assertEquals("#ff0000", sheet.getCell(11, 0).getCellStyle().getFillColor().getHtmlColor());
		assertEquals("#00ff00", sheet.getCell(11, 1).getCellStyle().getFillColor().getHtmlColor());
		assertEquals("#0000ff", sheet.getCell(11, 2).getCellStyle().getFillColor().getHtmlColor());

		//ensure cell style reusing
		assertTrue(sheet.getCell(27, 0).getCellStyle().equals(sheet.getCell(26, 0).getCellStyle()));
		assertTrue(sheet.getCell(28, 0).getCellStyle().equals(sheet.getCell(26, 0).getCellStyle()));
		assertTrue(sheet.getCell(28, 0).getCellStyle().getFont().equals(sheet.getCell(26, 0).getCellStyle().getFont()));
		
		//fill pattern
		assertEquals(FillPattern.SOLID_FOREGROUND, sheet.getCell(37, 1).getCellStyle().getFillPattern());
		assertEquals(FillPattern.ALT_BARS, sheet.getCell(37, 2).getCellStyle().getFillPattern());
		
		SSheet protectedSheet = book.getSheetByName("sheet-protection");
		assertEquals(true, protectedSheet.getCell(0, 0).getCellStyle().isLocked());
		assertEquals(false, protectedSheet.getCell(1, 0).getCellStyle().isLocked());
	}

	
	protected void cellBorderTest(SBook book){
		SSheet sheet = book.getSheetByName("cell-border");
		assertEquals(BorderType.NONE, sheet.getCell(2, 0).getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, sheet.getCell(2, 1).getCellStyle().getBorderBottom());
		assertEquals(BorderType.THIN, sheet.getCell(2, 2).getCellStyle().getBorderTop());
		assertEquals(BorderType.THIN, sheet.getCell(2, 3).getCellStyle().getBorderLeft());
		assertEquals(BorderType.THIN, sheet.getCell(2, 4).getCellStyle().getBorderRight());
		
		assertEquals(BorderType.HAIR, sheet.getCell(4, 1).getCellStyle().getBorderBottom());
		assertEquals(BorderType.DOTTED, sheet.getCell(4, 2).getCellStyle().getBorderBottom());
		assertEquals(BorderType.DASHED, sheet.getCell(4, 3).getCellStyle().getBorderBottom());
		
		assertEquals("#ff0000", sheet.getCell(14, 1).getCellStyle().getBorderBottomColor().getHtmlColor());
		assertEquals("#0000ff", sheet.getCell(14, 1).getCellStyle().getBorderLeftColor().getHtmlColor());
		assertEquals("#0000ff", sheet.getCell(14, 1).getCellStyle().getBorderTopColor().getHtmlColor());
		assertEquals("#ff0000", sheet.getCell(14, 1).getCellStyle().getBorderRightColor().getHtmlColor());
	}

	
	protected void cellFontNameTest(SBook book){
		SSheet sheet = book.getSheetByName("Style");
		assertEquals("Arial", sheet.getCell(3, 0).getCellStyle().getFont().getName());
		assertEquals("Arial Black", sheet.getCell(3, 1).getCellStyle().getFont().getName());
		assertEquals("Calibri", sheet.getCell(3, 2).getCellStyle().getFont().getName());
	}
	
	
	protected void cellFontStyleTest(SBook book){
		SSheet sheet = book.getSheetByName("Style");
		assertEquals(SFont.Boldweight.BOLD, sheet.getCell(9, 0).getCellStyle().getFont().getBoldweight());
		assertTrue(sheet.getCell(9, 1).getCellStyle().getFont().isItalic());
		assertTrue(sheet.getCell(9, 2).getCellStyle().getFont().isStrikeout());
		assertEquals(SFont.Underline.SINGLE, sheet.getCell(9, 3).getCellStyle().getFont().getUnderline());
		assertEquals(SFont.Underline.DOUBLE, sheet.getCell(9, 4).getCellStyle().getFont().getUnderline());
		assertEquals(SFont.Underline.SINGLE_ACCOUNTING, sheet.getCell(9, 5).getCellStyle().getFont().getUnderline());
		assertEquals(SFont.Underline.DOUBLE_ACCOUNTING, sheet.getCell(9, 6).getCellStyle().getFont().getUnderline());
		assertEquals(SFont.Underline.NONE, sheet.getCell(9, 7).getCellStyle().getFont().getUnderline());
		
		//height
		assertEquals(8, sheet.getCell(6, 0).getCellStyle().getFont().getHeightPoints());
		assertEquals(72, sheet.getCell(6, 3).getCellStyle().getFont().getHeightPoints());
		
		//type offset
		assertEquals(TypeOffset.SUPER, sheet.getCell(32, 1).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.SUB, sheet.getCell(32, 2).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.NONE, sheet.getCell(32, 3).getCellStyle().getFont().getTypeOffset());
	}
	
	
	protected void cellFontColorTest(SBook book){
		SSheet sheet = book.getSheetByName("Style");
		assertEquals("#000000", sheet.getCell(0, 0).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#ff0000", sheet.getCell(1, 0).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#00ff00", sheet.getCell(1, 1).getCellStyle().getFont().getColor().getHtmlColor());
		assertEquals("#0000ff", sheet.getCell(1, 2).getCellStyle().getFont().getColor().getHtmlColor());
	}
	
	
	protected void rowTest(SBook book){
		SSheet sheet = book.getSheetByName("Style");
		assertEquals(28, sheet.getRow(0).getHeight());
		assertEquals(20, sheet.getRow(1).getHeight());
		//style
		SCellStyle rowStyle1 = sheet.getRow(34).getCellStyle();
		assertEquals("#0000ff",rowStyle1.getFont().getColor().getHtmlColor());
		assertEquals(12,rowStyle1.getFont().getHeightPoints());
		SCellStyle rowStyle2 = sheet.getRow(35).getCellStyle();
		assertEquals(true,rowStyle2.getFont().isItalic());
		assertEquals(14,rowStyle2.getFont().getHeightPoints());		
		
		SSheet rowSheet = book.getSheetByName("column-row");
		assertTrue(rowSheet.getRow(9).isHidden());
		assertTrue(rowSheet.getRow(10).isHidden());
	}

	/**
	 * Information technology ??Document description and processing languages ??
	 * Office Open XML File Formats ??Part 1: Fundamentals and Markup LanguageReference  
	 * 18.8.30 numFmt (Number Format) 
	 */
	
	protected void cellFormatTest(SBook book){
		SSheet sheet = book.getSheetByName("Format");
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


	protected void columnTest(SBook book){
		SSheet sheet = book.getSheetByName("column-row");
		//column style
		assertEquals(SFont.Boldweight.BOLD, sheet.getColumn(0).getCellStyle().getFont().getBoldweight());
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
	
	protected void lastChangedColumnTest(SBook book){
		SSheet sheet = book.getSheetByName("column-row");
		assertEquals(110, sheet.getColumn(13).getWidth());
	}
	
	
	protected void viewInfoTest(SBook book){
		SSheet sheet = book.getSheetByName("column-row");
		assertEquals(3, sheet.getViewInfo().getNumOfRowFreeze());
		assertEquals(1, sheet.getViewInfo().getNumOfColumnFreeze());
		
		//grid line display
		assertTrue(sheet.getViewInfo().isDisplayGridlines());
		assertFalse(book.getSheetByName("cell-border").getViewInfo().isDisplayGridlines());
	}

	
	protected void mergedTest(SBook book){
		SSheet sheet = book.getSheetByName("cell-border");
		assertEquals(4, sheet.getMergedRegions().size());
		assertEquals("C31:D33", sheet.getMergedRegions().get(0).getReferenceString());
		assertEquals("B31:B33", sheet.getMergedRegions().get(1).getReferenceString());
		assertEquals("E28:G28", sheet.getMergedRegions().get(2).getReferenceString());
		assertEquals("B28:C28", sheet.getMergedRegions().get(3).getReferenceString());
	}	
	
	protected void areaChart(SBook book){
		SSheet sheet = book.getSheetByName("Area");
		SChart areaChart = sheet.getChart(0);
		assertEquals(ChartType.AREA,areaChart.getType());
		
		//a chart locating in one column and one row test
		assertEquals(493, areaChart.getAnchor().getWidth());
		assertEquals(283, areaChart.getAnchor().getHeight());
		
		SGeneralChartData chartData = (SGeneralChartData)areaChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
		
		SChart area3dChart = sheet.getChart(1);
		assertEquals(ChartGrouping.STANDARD, area3dChart.getGrouping());
		assertEquals(ChartLegendPosition.BOTTOM, area3dChart.getLegendPosition());
	}
	
	protected void barChart(SBook book){
		SSheet sheet = book.getSheetByName("Bar");
		SChart barChart = sheet.getChart(0);
		
		assertEquals(ChartType.BAR,barChart.getType());
		
		assertEquals(480, barChart.getAnchor().getWidth());
		assertEquals(284, barChart.getAnchor().getHeight());
		assertEquals(25, barChart.getAnchor().getXOffset());
		assertEquals(7, barChart.getAnchor().getYOffset());
		
		assertEquals(BarDirection.HORIZONTAL, barChart.getBarDirection());
		assertEquals(ChartGrouping.CLUSTERED, barChart.getGrouping());
		assertEquals(false, barChart.isThreeD());
		assertEquals(ChartLegendPosition.RIGHT, barChart.getLegendPosition());
		
		//data
		SGeneralChartData chartData = (SGeneralChartData)barChart.getData();
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
		
		SChart barChart3D = sheet.getChart(1);
		assertEquals(true, barChart3D.isThreeD());
	}
	
	public void bubbleChart(SBook book){
		SSheet sheet = book.getSheetByName("Bubble");
		SChart bubbleChart = sheet.getChart(0);
		assertEquals(ChartType.BUBBLE, bubbleChart.getType());
		
		SGeneralChartData chartData = (SGeneralChartData)bubbleChart.getData();
		assertEquals(0, chartData.getNumOfCategory());
		assertEquals(2, chartData.getNumOfSeries());
		assertEquals("String Literal Title", chartData.getSeries(0).getName());
		//has x, y, and z values
		assertEquals(5, chartData.getSeries(0).getNumOfXValue());
		assertEquals(5, chartData.getSeries(0).getNumOfYValue());
		assertEquals(5, chartData.getSeries(0).getNumOfZValue());
	}
	
	public void columnChart(SBook book){
		SSheet sheet = book.getSheetByName("Column");
		SChart columnChart = sheet.getChart(0);
		assertEquals(ChartType.COLUMN,columnChart.getType());
		assertEquals(BarDirection.VERTICAL, columnChart.getBarDirection());
		assertEquals(ChartLegendPosition.TOP, columnChart.getLegendPosition());
		
		SGeneralChartData chartData = (SGeneralChartData)columnChart.getData();
		assertEquals(4, chartData.getNumOfCategory());
		
		SChart column3dChart = sheet.getChart(1);
		assertEquals(ChartGrouping.STACKED, column3dChart.getGrouping());
	}
	
	public void doughnutChart(SBook book){
		SSheet sheet = book.getSheetByName("Doughnut");
		SChart doughnutChart = sheet.getChart(0);
		assertEquals(ChartType.DOUGHNUT, doughnutChart.getType());
		
		SGeneralChartData chartData = (SGeneralChartData)doughnutChart.getData();
		assertEquals(8, chartData.getNumOfCategory());
	}
	
	public void lineChart(SBook book){
		SSheet sheet = book.getSheetByName("Line");
		SChart lineChart = sheet.getChart(0);
		assertEquals(ChartType.LINE, lineChart.getType());
		SGeneralChartData chartData = (SGeneralChartData)lineChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
		
		SChart line3dChart = sheet.getChart(1);
		assertEquals(true, line3dChart.isThreeD());
		chartData = (SGeneralChartData)line3dChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
	}
	
	public void pieChart(SBook book){
		SSheet sheet = book.getSheetByName("Pie");
		SChart pieChart = sheet.getChart(0);
		assertEquals(ChartType.PIE, pieChart.getType());
		assertEquals(null,pieChart.getTitle());
		SGeneralChartData chartData = (SGeneralChartData)pieChart.getData();
		assertEquals(1, chartData.getNumOfSeries());
		
		SChart pie3dChart = sheet.getChart(1);
		assertEquals(ChartType.PIE, pie3dChart.getType());
		assertEquals(true, pie3dChart.isThreeD());
	}
	
	public void scatterChart(SBook book){
		SSheet sheet = book.getSheetByName("Scatter");
		SChart scatterChart = sheet.getChart(0);
		assertEquals(ChartType.SCATTER, scatterChart.getType());
		
		SGeneralChartData chartData = (SGeneralChartData)scatterChart.getData();
		assertEquals(3, chartData.getNumOfSeries());
		assertEquals("Internet Explorer", chartData.getSeries(0).getName());
		assertEquals(0.3427, chartData.getSeries(0).getYValue(0));
		assertEquals(0.327, chartData.getSeries(0).getYValue(1));
		assertEquals(0.3168, chartData.getSeries(0).getYValue(2));
		//has X and Y values
		assertEquals(8, chartData.getSeries(0).getNumOfXValue());
		assertEquals(8, chartData.getSeries(0).getNumOfYValue());
		assertEquals(0, chartData.getSeries(0).getNumOfZValue());
	}


	protected void picture(SBook book) {
		SSheet sheet1 = book.getSheet(0);
		assertEquals(2,sheet1.getPictures().size());
		SPicture zkLogo = sheet1.getPicture(0);
		assertEquals(Format.PNG, zkLogo.getFormat());
		assertEquals(450, zkLogo.getAnchor().getWidth());
		assertEquals(320, zkLogo.getAnchor().getHeight());
		
		SPicture zssBanner = sheet1.getPicture(1);
		assertEquals(Format.PNG, zssBanner.getFormat());
		assertEquals(275, zssBanner.getAnchor().getWidth());
		assertEquals(75, zssBanner.getAnchor().getHeight());

	}


	protected void validation(SBook book) {
		SSheet validationSheet = book.getSheetByName("Validation");
		assertEquals(7, validationSheet.getDataValidations().size());
		
		SDataValidation noValidation  = validationSheet.getDataValidation(0, 1);
		assertNull(noValidation);
	
		SDataValidation one2Ten  = validationSheet.getDataValidation(1, 1);
		assertEquals(ErrorStyle.STOP, one2Ten.getErrorStyle());
		assertEquals(OperatorType.BETWEEN, one2Ten.getOperatorType());
		assertEquals(ValidationType.INTEGER, one2Ten.getValidationType());
		//error box
		assertTrue(one2Ten.isShowErrorBox());
		assertEquals("Sorry", one2Ten.getErrorBoxTitle());
		assertEquals("1 - 10", one2Ten.getErrorBoxText());
		//prompt box
		assertTrue(one2Ten.isShowPromptBox());
		assertEquals("Notice", one2Ten.getPromptBoxTitle());
		assertEquals("valid between 1 to 10", one2Ten.getPromptBoxText());
		
		assertEquals(false, one2Ten.isShowDropDownArrow());
		assertEquals(true, one2Ten.isEmptyCellAllowed());
		assertEquals(true, one2Ten.isShowErrorBox());
		assertEquals(true, one2Ten.isShowPromptBox());
		
		SDataValidation fourGrades  = validationSheet.getDataValidation(2, 1);
		assertEquals(ValidationType.LIST, fourGrades.getValidationType());
		assertEquals(ErrorStyle.WARNING, fourGrades.getErrorStyle());
		assertEquals("$C$3:$F$3", fourGrades.getValue1Formula());
		assertEquals(4, fourGrades.getNumOfValue1());
		assertEquals(0, fourGrades.getNumOfValue2());
		assertEquals("A", fourGrades.getValue1(0).toString());
		assertEquals("B", fourGrades.getValue1(1).toString());
		assertEquals("C", fourGrades.getValue1(2).toString());
		assertEquals("D", fourGrades.getValue1(3).toString());
		assertEquals(false, fourGrades.isShowDropDownArrow());
		
		SDataValidation dayAfter2014  = validationSheet.getDataValidation(3, 1);
		assertEquals(ErrorStyle.INFO, dayAfter2014.getErrorStyle());
		
		SDataValidation lengthEquals10  = validationSheet.getDataValidation(4, 1);
		assertEquals(ErrorStyle.STOP, lengthEquals10.getErrorStyle());
		
		SDataValidation limitedColors  = validationSheet.getDataValidation(5, 1);
		assertEquals(ValidationType.LIST, limitedColors.getValidationType());
		assertEquals("\"red, blue, green\"", limitedColors.getValue1Formula());
		assertEquals(true, limitedColors.isShowDropDownArrow());
		
		SDataValidation custom  = validationSheet.getDataValidation(6, 1);
		assertEquals(ValidationType.FORMULA, custom.getValidationType());
		
		SDataValidation decimalRange  = validationSheet.getDataValidation(7, 1);
		assertEquals(ValidationType.DECIMAL, decimalRange.getValidationType());
	}


	protected void autoFilter(SBook book) {
		SAutoFilter filter1 = book.getSheetByName("1 column").getAutoFilter();
		assertEquals("B1:D10", filter1.getRegion().getReferenceString());
		assertEquals(1, filter1.getFilterColumns().size());
		assertEquals(FilterOp.VALUES, filter1.getFilterColumn(1, false).getOperator());
		assertEquals(1, filter1.getFilterColumn(1, false).getFilters().size());
		assertEquals(1, filter1.getFilterColumn(1, false).getCriteria1().size());
		assertTrue(filter1.getFilterColumn(1, false).getCriteria1().contains("Davolio"));
		
		
		SAutoFilter filter2 = book.getSheetByName("2 columns").getAutoFilter();
		assertEquals("A1:C21", filter2.getRegion().getReferenceString());
		assertEquals(2, filter2.getFilterColumns().size());
		
		NFilterColumn firstFilterColumn = filter2.getFilterColumn(0, false);
		assertEquals(FilterOp.VALUES, firstFilterColumn.getOperator());
		assertEquals(3, firstFilterColumn.getFilters().size());
		assertEquals(3, firstFilterColumn.getCriteria1().size());
		assertTrue(firstFilterColumn.getCriteria1().contains("XL"));
		assertTrue(firstFilterColumn.getCriteria1().contains("XXL"));
		assertTrue(firstFilterColumn.getCriteria1().contains("XXXL"));
		
		NFilterColumn secondFilterColumn = filter2.getFilterColumn(1, false);
		assertEquals(2, secondFilterColumn.getCriteria1().size());
		assertTrue(secondFilterColumn.getCriteria1().contains("Blue"));
		assertTrue(secondFilterColumn.getCriteria1().contains("Black"));
		
		SAutoFilter noCriteriaFilter = book.getSheetByName("no criteria").getAutoFilter();
		assertEquals("B2:D11", noCriteriaFilter.getRegion().getReferenceString());
		assertEquals(0, noCriteriaFilter.getFilterColumns().size());
	}
	
}

