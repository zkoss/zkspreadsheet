package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.model.Book;

@Ignore
public class ChartPictureIgnoredTest extends ChartPictureTestBase {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN,-1));
	}
	
	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testDeleteChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testDeleteChart(book);
	}
	
	// Unsupported
	@Test
	public void testAddBubbleChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddBubbleChart(book);
	}
	
	// Null pointer when select only one column or row or single cell
	@Test
	public void testAddScatterChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddScatterChart(book);
	}
	
	// Not Implement
	@Test
	public void testAddRadarChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddRadarChart(book);
	}
	
	// unsupported
	@Test
	public void testAddStockChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddStockChart(book);
	}
	
	// unsupported
	@Test
	public void testAddSurfaceChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddSurfaceChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddBarChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddBarChart(book);
	}
	
	// Not support chart operation in 2003
	@Test 
	public void testAddLineChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddLineChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddAreaChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddAreaChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddColumnChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddColumnChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddBubbleChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddBubbleChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddPieChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddPieChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddRadarChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddRadarChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddStockChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddStockChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddScatterChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddScatterChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddDoughnutChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddDoughnutChart(book);
	}
	
	// Not support chart operation in 2003
	@Test
	public void testAddSurfaceChart2003() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xls");
		testAddSurfaceChart(book);
	}
	
	
	@Test
	public void testDeletePicture2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testDeletePicture(book);
	}
	
	
	@Test
	public void testMovePicture2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testMovePicture(book);
	}

}
