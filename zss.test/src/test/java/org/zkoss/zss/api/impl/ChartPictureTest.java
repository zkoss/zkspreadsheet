package org.zkoss.zss.api.impl;

import static org.junit.Assert.*;

import java.io.*;
import java.util.Locale;

import org.junit.*;
import org.zkoss.zss.*;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.api.model.Chart.Type;
import org.zkoss.zss.ngapi.*;
import org.zkoss.zss.ngapi.impl.imexp.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NChart.NChartGrouping;
import org.zkoss.zss.ngmodel.NChart.NChartLegendPosition;
import org.zkoss.zss.ngmodel.NChart.NChartType;
import org.zkoss.zss.ngmodel.chart.NGeneralChartData;
import org.zkoss.zssex.api.ChartDataUtil;

public class ChartPictureTest extends ChartPictureTestBase {

	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
	
	@Test
	public void testMovePicture2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testMovePicture(book);
	}
	
	@Test
	public void testDeletePicture2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testDeletePicture(book);
	}

	
	@Test
	public void testAddPicture2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testAddPicture(book);
	}
	
	@Test
	public void testAddPicture2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testAddPicture(book);
	}
	
	@Test
	public void testDeleteChart2007() throws IOException {
		Book book = Util.loadBook(this,"book/insert-charts.xlsx");
		testDeleteChart(book);
	}
	
	@Test
	public void testAddBarChart2007() throws IOException {
		Book book = Util.loadBook(this,"book/insert-charts.xlsx");
		testAddBarChart(book);
	}
	
	@Test
	public void testAddLineChart2007() throws IOException {
		Book book = Util.loadBook(this,"book/insert-charts.xlsx");
		testAddLineChart(book);
	}
	
	@Test
	public void testAddAreaChart2007() throws IOException {
		Book book = Util.loadBook(this,"book/insert-charts.xlsx");
		testAddAreaChart(book);
	}
	
	@Test
	public void testAddColumnChart2007() throws IOException {
		Book book = Util.loadBook(this,"book/insert-charts.xlsx");
		testAddColumnChart(book);
	}
	
	@Test
	public void testAddPieChart2007() throws IOException {
		Book book = Util.loadBook(this,"book/insert-charts.xlsx");
		testAddPieChart(book);
	}
	
	@Test
	public void testAddDoughnutChart2007() throws IOException {
		Book book = Util.loadBook(this,"book/insert-charts.xlsx");
		testAddDoughnutChart(book);
	}
	
	/**
	 * For data area, row count is less or equal than column count.
	 */
	@Test
	public void addBarChart(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(2, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A4:D7"), Type.BAR, anchor);
		//FIXME how to know 3D
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.BAR, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);

		assertEquals(1, sheet.getCharts().size());
		NGeneralChartData nChartData = (NGeneralChartData)chart.getData();
		assertEquals(3, nChartData.getNumOfCategory());
		assertEquals("Internet Explorer", nChartData.getCategory(0));
		assertEquals("Chrome", nChartData.getCategory(1));
		assertEquals("Firefox", nChartData.getCategory(2));
		assertEquals(3, nChartData.getNumOfSeries());
		assertEquals("January 2012", nChartData.getSeries(0).getName()); 
		assertEquals(0.3427, nChartData.getSeries(0).getValue(0));
		assertEquals(0.2599, nChartData.getSeries(0).getValue(1));
		assertEquals(0.2268, nChartData.getSeries(0).getValue(2));
		assertEquals("February 2012", nChartData.getSeries(1).getName());
		assertEquals(0.327, nChartData.getSeries(1).getValue(0));
		assertEquals(0.2724, nChartData.getSeries(1).getValue(1));
		assertEquals(0.2276, nChartData.getSeries(1).getValue(2));
		assertEquals("March 2012", nChartData.getSeries(2).getName());
		assertEquals(0.3168, nChartData.getSeries(2).getValue(0));
		assertEquals(0.2809, nChartData.getSeries(2).getValue(1));
		assertEquals(0.2273, nChartData.getSeries(2).getValue(2));

		write(book,  org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type.XLSX); //human checking
	}
	
	/**
	 * For data area, row count is larger than column count.
	 */
	@Test
	public void addColumnChart(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(2, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A4:D9"), Type.COLUMN, anchor);
		//FIXME how to know 3D
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.COLUMN, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);

		assertEquals(1, sheet.getCharts().size());
		NGeneralChartData nChartData = (NGeneralChartData)chart.getData();
		assertEquals(5, nChartData.getNumOfCategory());
		assertEquals("January 2012", nChartData.getCategory(0));
		assertEquals("February 2012", nChartData.getCategory(1));
		assertEquals("March 2012", nChartData.getCategory(2));
		assertEquals(3, nChartData.getNumOfSeries());
		assertEquals("Internet Explorer", nChartData.getSeries(0).getName());
		assertEquals(0.3427, nChartData.getSeries(0).getValue(0));
		assertEquals(0.327, nChartData.getSeries(0).getValue(1));
		assertEquals(0.3168, nChartData.getSeries(0).getValue(2));
		
		assertEquals("Chrome", nChartData.getSeries(1).getName());
		assertEquals(0.2599, nChartData.getSeries(1).getValue(0));
		assertEquals(0.2724, nChartData.getSeries(1).getValue(1));
		assertEquals(0.2809, nChartData.getSeries(1).getValue(2));
		
		assertEquals("Firefox", nChartData.getSeries(2).getName());
		assertEquals(0.2268, nChartData.getSeries(2).getValue(0));
		assertEquals(0.2276, nChartData.getSeries(2).getValue(1));
		assertEquals(0.2273, nChartData.getSeries(2).getValue(2));

//		write(book,  org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type.XLSX); //human checking
	}
	
	@Test
	public void addColumnChartMissingCategory(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(2, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("B4:D9"), Type.COLUMN, anchor);
		//FIXME how to know 3D
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.COLUMN, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);
		assertEquals(1, sheet.getCharts().size());

		NGeneralChartData nChartData = (NGeneralChartData)chart.getData();
		assertEquals(0, nChartData.getNumOfCategory());
		assertEquals(3, nChartData.getNumOfSeries());
		assertEquals(0.3427, nChartData.getSeries(0).getValue(0));
		assertEquals(0.327, nChartData.getSeries(0).getValue(1));
		assertEquals(0.3168, nChartData.getSeries(0).getValue(2));
		
		assertEquals(0.2599, nChartData.getSeries(1).getValue(0));
		assertEquals(0.2724, nChartData.getSeries(1).getValue(1));
		assertEquals(0.2809, nChartData.getSeries(1).getValue(2));
		
		assertEquals(0.2268, nChartData.getSeries(2).getValue(0));
		assertEquals(0.2276, nChartData.getSeries(2).getValue(1));
		assertEquals(0.2273, nChartData.getSeries(2).getValue(2));

//		write(book,  org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type.XLSX); //human checking
	}
	
	@Test
	public void addColumnChartMissingSeriesName(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(2, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A5:D9"), Type.COLUMN, anchor);
		//FIXME how to know 3D
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.COLUMN, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);
		assertEquals(1, sheet.getCharts().size());

		NGeneralChartData nChartData = (NGeneralChartData)chart.getData();
		assertEquals(5, nChartData.getNumOfCategory());
		assertEquals("January 2012", nChartData.getCategory(0));
		assertEquals("February 2012", nChartData.getCategory(1));
		assertEquals("March 2012", nChartData.getCategory(2));
		assertEquals(3, nChartData.getNumOfSeries());
		assertEquals(0.3427, nChartData.getSeries(0).getValue(0));
		assertEquals(0.327, nChartData.getSeries(0).getValue(1));
		assertEquals(0.3168, nChartData.getSeries(0).getValue(2));
		
		assertEquals(0.2599, nChartData.getSeries(1).getValue(0));
		assertEquals(0.2724, nChartData.getSeries(1).getValue(1));
		assertEquals(0.2809, nChartData.getSeries(1).getValue(2));
		
		assertEquals(0.2268, nChartData.getSeries(2).getValue(0));
		assertEquals(0.2276, nChartData.getSeries(2).getValue(1));
		assertEquals(0.2273, nChartData.getSeries(2).getValue(2));

//		write(book,  org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type.XLSX); //human checking
	}
		
	/**
	 * To test fill chart data with x & y values. 
	 */
	@Test
	public void addScatterChart(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(2, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A4:D9"), Type.SCATTER, anchor);
		//FIXME how to know 3D
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.SCATTER, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);
		assertEquals(1, sheet.getCharts().size());
		
		NGeneralChartData nChartData = (NGeneralChartData)chart.getData();
		assertEquals(0, nChartData.getNumOfCategory());
		assertEquals(3, nChartData.getNumOfSeries());
		assertEquals("January 2012", nChartData.getSeries(0).getValue(0));
		assertEquals("February 2012", nChartData.getSeries(0).getValue(1));
		assertEquals("March 2012", nChartData.getSeries(0).getValue(2));
		
		assertEquals(0.3427, nChartData.getSeries(0).getYValue(0));
		assertEquals(0.327, nChartData.getSeries(0).getYValue(1));
		assertEquals(0.3168, nChartData.getSeries(0).getYValue(2));
		
		assertEquals(0.2599, nChartData.getSeries(1).getYValue(0));
		assertEquals(0.2724, nChartData.getSeries(1).getYValue(1));
		assertEquals(0.2809, nChartData.getSeries(1).getYValue(2));
		
		assertEquals(0.2268, nChartData.getSeries(2).getYValue(0));
		assertEquals(0.2276, nChartData.getSeries(2).getYValue(1));
		assertEquals(0.2273, nChartData.getSeries(2).getYValue(2));

//		write(book,  org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type.XLSX); //human checking
	}
	
	/**
	 * different data source. missing series name.
	 */
	@Test
	public void addScatterChart2(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(30, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A30:B39"), Type.SCATTER, anchor);
		//FIXME how to know 3D
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.SCATTER, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);

		assertEquals(1, sheet.getCharts().size());
		NGeneralChartData nChartData = (NGeneralChartData)chart.getData();
		assertEquals(0, nChartData.getNumOfCategory());
		assertEquals(1, nChartData.getNumOfSeries());

		assertEquals("Particulate", nChartData.getSeries(0).getName());
		
		assertEquals(4.1, nChartData.getSeries(0).getValue(0));
		assertEquals(4.3, nChartData.getSeries(0).getValue(1));
		assertEquals(5.7, nChartData.getSeries(0).getValue(2));
		
		assertEquals(122.0, nChartData.getSeries(0).getYValue(0));
		assertEquals(117.0, nChartData.getSeries(0).getYValue(1));
		assertEquals(112.0, nChartData.getSeries(0).getYValue(2));
		
//		write(book,  org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type.XLSX); //human checking
		
	}
	
	/**
	 * column count equal row count.
	 */
	@Test
	public void addScatterChart3(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(30, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A30:B32"), Type.SCATTER, anchor);
		//FIXME how to know 3D
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.SCATTER, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);

		assertEquals(1, sheet.getCharts().size());
		NGeneralChartData nChartData = (NGeneralChartData)chart.getData();
		assertEquals(0, nChartData.getNumOfCategory());
		assertEquals(2, nChartData.getNumOfSeries());

		assertNull(nChartData.getSeries(0).getName());
		assertNull(nChartData.getSeries(1).getName());
	}
	
	@Test
	public void deleteChart(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(2, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A4:D9"), Type.COLUMN, anchor);
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.COLUMN, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);
		assertEquals(1, sheet.getCharts().size());
		
		NRanges.range(sheet).deleteChart(chart);
		assertEquals(0, sheet.getCharts().size());
	}
	
	@Test
	public void moveChart(){
		NBook book = loadBook(this.getClass().getResourceAsStream("book/insert-charts.xlsx"));
		NSheet sheet = book.getSheetByName("chart-image");
		
		assertEquals(0, sheet.getCharts().size());
		NViewAnchor anchor = new NViewAnchor(2, 5, 600, 400);
		ChartData chartData =  ChartDataUtil.getChartData(sheet, new AreaRef("A4:D9"), Type.COLUMN, anchor);
		NChart chart = NRanges.range(sheet).addChart(anchor, chartData.getNative(), NChartType.COLUMN, NChartGrouping.STANDARD, NChartLegendPosition.RIGHT);
		assertEquals(1, sheet.getCharts().size());
		assertEquals(2, chart.getAnchor().getRowIndex());
		assertEquals(5, chart.getAnchor().getColumnIndex());
		
		NRanges.range(sheet).moveChart(chart, new NViewAnchor(5, 8, 700, 500));
		assertEquals(5, chart.getAnchor().getRowIndex());
		assertEquals(8, chart.getAnchor().getColumnIndex());
	}
	
	
	//TODO use ImExpTestUtil.loadBook()
	public static NBook loadBook(InputStream is) {
		NImporter importer = new ExcelImportFactory().createImporter();
		NBook book = null;
		try {
			book = importer.imports(is, "book");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return book;
	}
	static public String DEFAULT_EXPORT_TARGET_PATH = "./target/";
	static public String DEFAULT_EXPORT_FILE_NAME_XLSX = "exported.xlsx";
	static public String DEFAULT_EXPORT_FILE_NAME_XLS = "exported.xls";
	
	public static File write(NBook book, org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type type) {
		if (type.equals(ExcelExportFactory.Type.XLSX)){
			return writeBookToFile(book, new File(DEFAULT_EXPORT_TARGET_PATH + DEFAULT_EXPORT_FILE_NAME_XLSX), type);
		}else{
			return writeBookToFile(book, new File(DEFAULT_EXPORT_TARGET_PATH + DEFAULT_EXPORT_FILE_NAME_XLS), type);
		}
	}
	
	public static File writeBookToFile(NBook book, File outFile, org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory.Type type) {
		try {
			outFile = new File(DEFAULT_EXPORT_TARGET_PATH + outFile.getName());
			outFile.createNewFile();
			NExporter exporter = new ExcelExportFactory(type).createExporter();
			exporter.export(book, outFile);
		} catch (Exception e) {
			e.printStackTrace();
			fail(e.toString());
		}
		return outFile;
	}
}
