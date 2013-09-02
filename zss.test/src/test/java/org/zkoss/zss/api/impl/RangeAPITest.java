package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.io.InputStream;

import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Hyperlink;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.impl.CellStyleImpl;
import org.zkoss.zssex.api.ChartDataUtil;

public class RangeAPITest {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Ignore("Not support chart operation in 2003")
	@Test
	public void testInsertDeleteChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testInsertDeleteChart(book, "book/test.xls");
	}
	
	@Test
	public void testInsertDeleteChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testInsertDeleteChart(book, "book/test.xlsx");
	}
	
	@Test
	public void testBasicAPI2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testBasicAPI(book);
	}
	
	@Test
	public void testBasicAPI2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testBasicAPI(book);
	}
	
	@Test
	public void testHyperLink2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testHyperLink(book, "book/test.xls");
	}
	
	@Test
	public void testHyperLink2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testHyperLink(book, "book/test.xlsx");
	}
	
	/**
	 * Bubble: unsupported
	 * Radar: Not Implement
	 * Stock: unsupported
	 * Surface: unsupported
	 * Scatter: Null pointer when select only one column or row or single cell
	 */
	private void testInsertDeleteChart(Book workbook, String outFileName) throws IOException {
		
		Sheet sheet = workbook.getSheet("chart-image");
		
		// Insert chart 1, 2 (LINE)
		ChartData cd1 = ChartDataUtil.getChartData(sheet, new AreaRef(4,1,14,1), Chart.Type.LINE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "A1"), cd1, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		
		ChartData cd2 = ChartDataUtil.getChartData(sheet, new AreaRef(4,2,14,2), Chart.Type.LINE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "I1"), cd2, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);

		// BAR
		ChartData cd3 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.BAR);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q1"), cd3, Chart.Type.BAR, Grouping.STANDARD, LegendPosition.TOP);
		
		// AREA
		ChartData cd4 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.AREA);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q10"), cd4, Chart.Type.AREA, Grouping.STANDARD, LegendPosition.TOP);
		
		// BUBBLE (Unsupported)
		// ChartData cd5 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.BUBBLE);
		// SheetOperationUtil.addChart(Ranges.range(sheet, "Q20"), cd5, Chart.Type.BUBBLE, Grouping.STANDARD, LegendPosition.TOP);
		
		// COLUMN
		ChartData cd6 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.COLUMN);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q30"), cd6, Chart.Type.COLUMN, Grouping.STANDARD, LegendPosition.TOP);
				
		// PIE
		ChartData cd7 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.PIE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q40"), cd7, Chart.Type.PIE, Grouping.STANDARD, LegendPosition.TOP);
		
		// RADAR (NOT IMPLEMENT)
		// ChartData cd8 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.RADAR);
		// SheetOperationUtil.addChart(Ranges.range(sheet, "Q50"), cd8, Chart.Type.RADAR, Grouping.STANDARD, LegendPosition.TOP);
		
		// SCATTER (NULL POINTER)
		// ChartData cd9 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.SCATTER);
		// SheetOperationUtil.addChart(Ranges.range(sheet, "Q60"), cd9, Chart.Type.SCATTER, Grouping.STANDARD, LegendPosition.TOP);
		
		// STOCK (Unsupported)
		// ChartData cd10 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.STOCK);
		// SheetOperationUtil.addChart(Ranges.range(sheet, "Q70"), cd10, Chart.Type.STOCK, Grouping.STANDARD, LegendPosition.TOP);
								
		// SURFACE (Unsupported)
		// ChartData cd11 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.SURFACE);
		// SheetOperationUtil.addChart(Ranges.range(sheet, "Q80"), cd11, Chart.Type.SURFACE, Grouping.STANDARD, LegendPosition.TOP);
		
		// DOUGHNUT
		ChartData cd11 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.DOUGHNUT);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q90"), cd11, Chart.Type.DOUGHNUT, Grouping.STANDARD, LegendPosition.TOP);
		
		Util.export(workbook, outFileName);
		
		workbook = Util.loadBook(outFileName);
		sheet = workbook.getSheet("chart-image");
		
		// FIXME
		// MUST initDrawing before sheet.getCharts
//		DrawingManagerImpl d = new DrawingManagerImpl(sheet.getPoiSheet());
		// Not work
		// System.out.println(sheet.getCharts().size());
//		d.getCharts();
//		for(Chart chart: ) {
//			SheetAnchor anchor = chart.getAnchor();
//			int row1 = anchor.getRow();
//			int col1 = anchor.getColumn();
//			int row2 = anchor.getLastRow();
//			int col2 = anchor.getLastColumn();
//			SheetOperationUtil.deleteChart(Ranges.range(sheet, row1, col1, row2, col2), chart);
//		}
		
		Util.export(workbook, outFileName);
	}
	
	
	
	/**
	 * getColumn()
	 * getColumnCount()
	 * getRow()
	 * getLastColumn()
	 * getLastRow()
	 * getRowCount()
	 * toRowRange()
	 * toColumnRange()
	 * isWholeRow()
	 * isWholeColumn()
	 * clearContents()
	 * clearStyle()
	 */
	private void testBasicAPI(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertEquals(0, rangeA5F18.getColumn());
		assertEquals(6, rangeA5F18.getColumnCount());
		assertEquals(4, rangeA5F18.getRow());
		assertEquals(5, rangeA5F18.getLastColumn());
		assertEquals(17, rangeA5F18.getLastRow());
		assertEquals(14, rangeA5F18.getRowCount());
		assertTrue(!rangeA5F18.isWholeRow());
		assertTrue(!rangeA5F18.isWholeColumn());
		
		Range rowRange = rangeA5F18.toRowRange();
		assertTrue(rowRange.isWholeRow());
		
		Range colRange = rangeA5F18.toColumnRange();
		assertTrue(colRange.isWholeColumn());
		
		// clearContents()
		Range rangeA1 = Ranges.range(sheet, "A1");
		rangeA1.setCellEditText("ZK");
		assertEquals("ZK", rangeA1.getCellEditText());
		CellOperationUtil.applyFontColor(rangeA1, "#ff0000");
		CellOperationUtil.applyFontStrikeout(rangeA1, true);
		assertEquals("#ff0000", rangeA1.getCellStyle().getFont().getColor().getHtmlColor());
		assertTrue(rangeA1.getCellStyle().getFont().isStrikeout());
		rangeA1.clearContents();
		assertEquals(CellType.BLANK, rangeA1.getCellData().getType());
		assertEquals("#ff0000", rangeA1.getCellStyle().getFont().getColor().getHtmlColor());
		assertTrue(rangeA1.getCellStyle().getFont().isStrikeout());
		
		// clearStyles()
		rangeA1.clearStyles();
		assertEquals("#000000", rangeA1.getCellStyle().getFont().getColor().getHtmlColor());
		assertTrue(!rangeA1.getCellStyle().getFont().isStrikeout());
	}
	
	/**
	 * Test HyperLink
	 */
	private void testHyperLink(Book workbook, String outFileName) throws IOException {

		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1 = Ranges.range(sheet, "A1");
		rA1.setCellHyperlink(HyperlinkType.URL, "http://www.zkoss.org", "ZK");
		
		Range rA2 = Ranges.range(sheet, "A2");
		// set "info@zkoss.org" will cause illegal argument exception
		// set "mailto:info@zkoss.org" is correct
		rA2.setCellHyperlink(HyperlinkType.EMAIL, "mailto:info@zkoss.org", "ZK");
		
		rA2 = Ranges.range(sheet, "A2");
		Hyperlink linkEmail = rA2.getCellHyperlink();
		assertTrue(linkEmail != null);
		assertEquals("mailto:info@zkoss.org", linkEmail.getAddress());
		assertEquals(HyperlinkType.EMAIL, linkEmail.getType());
		
		Range rA3 = Ranges.range(sheet, "A3");
		rA3.setCellHyperlink(HyperlinkType.FILE, "test.xlsx", "ZK");
		
		Range rA4 = Ranges.range(sheet, "A4");
		rA4.setCellHyperlink(HyperlinkType.DOCUMENT, "test.xlsx", "ZK");
		
		Util.export(workbook, outFileName);
		
		// reload
		workbook = Util.loadBook(outFileName);
		sheet = workbook.getSheet("Sheet1");
		
		rA1 = Ranges.range(sheet, "A1");
		Hyperlink linkURL = rA1.getCellHyperlink();
		assertTrue(linkURL != null);
		assertEquals("http://www.zkoss.org", linkURL.getAddress());
		assertEquals(HyperlinkType.URL, linkURL.getType());
		
		rA2 = Ranges.range(sheet, "A2");
		linkEmail = rA2.getCellHyperlink();
		assertTrue(linkEmail != null);
		assertEquals("mailto:info@zkoss.org", linkEmail.getAddress());
		assertEquals(HyperlinkType.EMAIL, linkEmail.getType());
		
		rA3 = Ranges.range(sheet, "A3");
		Hyperlink linkFile = rA3.getCellHyperlink();
		assertTrue(linkFile != null);
		assertEquals("test.xlsx", linkFile.getAddress());
		assertEquals(HyperlinkType.FILE, linkFile.getType());
		
		rA4 = Ranges.range(sheet, "A4");
		Hyperlink linkDoc = rA4.getCellHyperlink();
		assertTrue(linkDoc != null);
		assertEquals("test.xlsx", linkDoc.getAddress());
		assertEquals(HyperlinkType.DOCUMENT, linkDoc.getType());
		
		Util.export(workbook, outFileName);
		
		rA1.setCellEditText("URL");
		//assertTrue(rA1.getCellHyperlink() == null);
		assertEquals("URL", rA1.getCellEditText());
		
		rA2.setCellEditText("Email");
		//assertTrue(rA2.getCellHyperlink() == null);
		assertEquals("Email", rA2.getCellEditText());
		
		rA3.setCellEditText("File");
		//assertTrue(rA3.getCellHyperlink() == null);
		assertEquals("File", rA3.getCellEditText());
		
		rA4.setCellEditText("Document");
		//assertTrue(rA4.getCellHyperlink() == null);
		assertEquals("Document", rA4.getCellEditText());
		
		Util.export(workbook, outFileName);
		
	}
	
}
