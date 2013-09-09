package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.model.Book;

public class ChartPictureTest extends ChartPictureTestBase {

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
	
	@Test
	public void testMovePicture2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testMovePicture(book);
	}
	
	@Test
	public void testDeletePicture2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testDeletePicture(book);
	}

	
	@Test
	public void testAddPicture2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testAddPicture(book);
	}
	
	@Test
	public void testAddPicture2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testAddPicture(book);
	}
	
	@Test
	public void testDeleteChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testDeleteChart(book);
	}
	
	@Test
	public void testAddBarChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddBarChart(book);
	}
	
	@Test
	public void testAddLineChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddLineChart(book);
	}
	
	@Test
	public void testAddAreaChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddAreaChart(book);
	}
	
	@Test
	public void testAddColumnChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddColumnChart(book);
	}
	
	@Test
	public void testAddPieChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddPieChart(book);
	}
	
	@Test
	public void testAddDoughnutChart2007() throws IOException {
		Book book = Util.loadBook("book/insert-charts.xlsx");
		testAddDoughnutChart(book);
	}
	
}
