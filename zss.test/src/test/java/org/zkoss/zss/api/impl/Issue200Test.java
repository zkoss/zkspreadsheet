package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.image.AImage;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.AssertUtil;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zssex.api.ChartDataUtil;

/**
 * ZSS-245.
 * ZSS-255.
 * ZSS-260. ZSS-261. ZSS-262. ZSS-263. ZSS-264. ZSS-265. ZSS-266. ZSS-267.
 * ZSS-275. ZSS-277. ZSS-280.
 * ZSS-290. ZSS-298.
 * ZSS-300. ZSS-301. ZSS-303. 
 * ZSS-315. ZSS-317. 
 * ZSS-326. 
 * ZSS-334. ZSS-341. ZSS-342.
 * ZSS-355.
 * ZSS-389.
 * ZSS-395.
 * ZSS-399.
 */
public class Issue200Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}
	
	@Test
	public void testZSS291() throws IOException {
		final String filename = "book/291-sort.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("cell");
		Ranges.range(sheet, "B4:B8").sort(false);

		assertEquals(1, Ranges.range(sheet, "B4").getCellData().getDoubleValue(), 1E-8);
		assertEquals(2, Ranges.range(sheet, "B5").getCellData().getDoubleValue(), 1E-8);
		assertEquals(3, Ranges.range(sheet, "B6").getCellData().getDoubleValue(), 1E-8);
		assertEquals(4, Ranges.range(sheet, "B7").getCellData().getDoubleValue(), 1E-8);
		assertEquals(5, Ranges.range(sheet, "B8").getCellData().getDoubleValue(), 1E-8);
	}
	
	@Ignore("ZSS-280")
	@Test
	public void testZSS280() throws IOException {
		final String filename = "book/280-autofilter.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("cell-data");
		Ranges.range(sheet, "A13").setCellEditText("1");
		Ranges.range(sheet, "A14").setCellEditText("2");
		Ranges.range(sheet, "A15").setCellEditText("3");
		Ranges.range(sheet, "A16").setCellEditText("4");
		Ranges.range(sheet, "A17").setCellEditText("5");
		SheetOperationUtil.applyAutoFilter(Ranges.range(sheet, "A1:A17"));
		
		AutoFilter af = sheet.getPoiSheet().getAutoFilter();
		if (af == null) { //no AutoFilter to apply 
			return;
		}
		final CellRangeAddress affectedArea = af.getRangeAddress();
		final int row1 = affectedArea.getFirstRow();
		final int col1 = affectedArea.getFirstColumn(); 
		final int row2 = affectedArea.getLastRow();
		final int col2 = affectedArea.getLastColumn();
		
		assertEquals(row1, 0);
		assertEquals(col1, 0);
		assertEquals(row2, 17);
		assertEquals(col2, 0);
	}
	
	@Ignore("ZSS-271")
	@Test
	public void testZSS271() throws IOException {
		final String filename = "book/271-engineering.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("formula-engineering");
		assertEquals("0.98", Ranges.range(sheet, "B3").getCellFormatText());
		assertEquals("0.33", Ranges.range(sheet, "B5").getCellFormatText());
		assertEquals("0.28", Ranges.range(sheet, "B7").getCellFormatText());
		assertEquals("0.15", Ranges.range(sheet, "B9").getCellFormatText());
		assertEquals("100", Ranges.range(sheet, "B11").getCellFormatText());
		assertEquals("00FB", Ranges.range(sheet, "B13").getCellFormatText());
		assertEquals("011", Ranges.range(sheet, "B15").getCellFormatText());
		assertEquals("3+4i", Ranges.range(sheet, "B17").getCellFormatText());
		assertEquals("1001", Ranges.range(sheet, "B19").getCellFormatText());
		assertEquals("0064", Ranges.range(sheet, "B21").getCellFormatText());
		assertEquals("072", Ranges.range(sheet, "B23").getCellFormatText());
		assertEquals("0", Ranges.range(sheet, "B25").getCellFormatText());
		assertEquals("0.71", Ranges.range(sheet, "B27").getCellFormatText());
		assertEquals("0.16", Ranges.range(sheet, "B29").getCellFormatText());
		assertEquals("1", Ranges.range(sheet, "B31").getCellFormatText());
		assertEquals("00001111", Ranges.range(sheet, "B33").getCellFormatText());
		assertEquals("165", Ranges.range(sheet, "B35").getCellFormatText());
		assertEquals("017", Ranges.range(sheet, "B37").getCellFormatText());
		assertEquals("13", Ranges.range(sheet, "B39").getCellFormatText());
		assertEquals("4", Ranges.range(sheet, "B41").getCellFormatText());
		assertEquals("0.93", Ranges.range(sheet, "B43").getCellFormatText());
		assertEquals("3-4i", Ranges.range(sheet, "B45").getCellFormatText());
		assertEquals("0.833730025131149-0.988897705762865i", Ranges.range(sheet, "B47").getCellFormatText());
		assertEquals("5+12i", Ranges.range(sheet, "B49").getCellFormatText());
		assertEquals("1.46869393991589+2.28735528717884i", Ranges.range(sheet, "B51").getCellFormatText());
		assertEquals("1.6094379124341+0.927295218001612i", Ranges.range(sheet, "B53").getCellFormatText());
		assertEquals("0.698970004336019+0.402719196273373i", Ranges.range(sheet, "B55").getCellFormatText());
		assertEquals("2.32192809506607+1.33780421255394i", Ranges.range(sheet, "B57").getCellFormatText());
		assertEquals("-46+9.00000000000001i", Ranges.range(sheet, "B59").getCellFormatText());
		assertEquals("27+11i", Ranges.range(sheet, "B61").getCellFormatText());
		assertEquals("6", Ranges.range(sheet, "B63").getCellFormatText());
		assertEquals("3.85373803791938-27.0168132580039i", Ranges.range(sheet, "B65").getCellFormatText());
		assertEquals("1.09868411346781+0.455089860562227i", Ranges.range(sheet, "B67").getCellFormatText());
		assertEquals("8+i", Ranges.range(sheet, "B69").getCellFormatText());
		assertEquals("8+i", Ranges.range(sheet, "B71").getCellFormatText());
		assertEquals("011", Ranges.range(sheet, "B73").getCellFormatText());
		assertEquals("44", Ranges.range(sheet, "B75").getCellFormatText());
		assertEquals("0040", Ranges.range(sheet, "B77").getCellFormatText());
	}
	
	
	@Ignore("ZSS-272")
	@Test
	public void testZSS272() throws IOException {
		final String filename = "book/272-conditionalFormatting.xls";
		Book workbook = Util.loadBook(this,filename);
		
		//shouldn't throw any exception
		((HSSFSheet)workbook.getSheetAt(0).getPoiSheet()).getDataValidations();
	}
	
	@Ignore("ZSS-270")
	@Test
	public void testZSS270() throws IOException {
		final String filename = "book/270-statistical.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("formula-statistical");
		 // AVEDEV
		assertEquals("1.02", Ranges.range(sheet, "B3").getCellFormatText());
		// AVERAGE
		assertEquals("11", Ranges.range(sheet, "B6").getCellFormatText());
		// AVERAGEA
		assertEquals("5.6", Ranges.range(sheet, "B8").getCellFormatText());
		// AVERAGEIF
		assertEquals("14000", Ranges.range(sheet, "B11").getCellFormatText());
		// AVERAGEIFS
		assertEquals("87.5", Ranges.range(sheet, "B14").getCellFormatText());
		// BETADIST
		assertEquals("0.69", Ranges.range(sheet, "B17").getCellFormatText());
		// BETAINV
		assertEquals("2", Ranges.range(sheet, "B19").getCellFormatText());
		// BIOMDIST
		assertEquals("0.21", Ranges.range(sheet, "B21").getCellFormatText());
		// CHIDIST
		assertEquals("0.05", Ranges.range(sheet, "B23").getCellFormatText());
		// CHIINV
		assertEquals("18.31", Ranges.range(sheet, "B25").getCellFormatText());
		// CHITEST
		assertEquals("0.000308", Ranges.range(sheet, "B27").getCellFormatText());
		// CONFIDENCE
		assertEquals("0.69", Ranges.range(sheet, "B35").getCellFormatText());
		// CORREL
		assertEquals("0.9971", Ranges.range(sheet, "B37").getCellFormatText());
		// COUNT
		assertEquals("3", Ranges.range(sheet, "B41").getCellFormatText());
		// COUNTA
		assertEquals("6", Ranges.range(sheet, "B43").getCellFormatText());
		// COUNTBLANK
		assertEquals("2", Ranges.range(sheet, "B45").getCellFormatText());
		// COUNTIF
		assertEquals("2", Ranges.range(sheet, "B47").getCellFormatText());
		// COVAR
		assertEquals("5.2", Ranges.range(sheet, "B50").getCellFormatText());
		// CRITBINOM
		assertEquals("4", Ranges.range(sheet, "B53").getCellFormatText());
		// DEVSQ
		assertEquals("48", Ranges.range(sheet, "B55").getCellFormatText());
		// EXPONDIST
		assertEquals("0.86", Ranges.range(sheet, "B57").getCellFormatText());
		// FDIST
		assertEquals("0.01", Ranges.range(sheet, "B59").getCellFormatText());
		// FINV
		assertEquals("15.21", Ranges.range(sheet, "B61").getCellFormatText());
		// FISHER
		assertEquals("0.97", Ranges.range(sheet, "B63").getCellFormatText());
		// FISHERINV
		assertEquals("0.75", Ranges.range(sheet, "B65").getCellFormatText());
		// FORECAST
		assertEquals("10.61", Ranges.range(sheet, "B67").getCellFormatText());
		// FREQUENCY
		assertEquals("1", Ranges.range(sheet, "B70").getCellFormatText());
		// FTEST
		assertEquals("0.65", Ranges.range(sheet, "B73").getCellFormatText());
		// GAMMADIST
		assertEquals("0.03", Ranges.range(sheet, "B76").getCellFormatText());
		// GAMMAINV
		assertEquals("10.00", Ranges.range(sheet, "B78").getCellFormatText());
		// GAMMALN
		assertEquals("1.79", Ranges.range(sheet, "B80").getCellFormatText());
		// GEOMEAN
		assertEquals("5.48", Ranges.range(sheet, "B82").getCellFormatText());
		// GROWTH
		assertEquals("320196.72", Ranges.range(sheet, "B85").getCellFormatText());
		// HARMEAN
		assertEquals("5.03", Ranges.range(sheet, "B89").getCellFormatText());
		// HYPGEOMDIST
		assertEquals("0.36", Ranges.range(sheet, "B92").getCellFormatText());
		// INTERCEPT
		assertEquals("0.05", Ranges.range(sheet, "B94").getCellFormatText());
		// KURT
		assertEquals("-0.15", Ranges.range(sheet, "B97").getCellFormatText());
		// LARGE
		assertEquals("5", Ranges.range(sheet, "B100").getCellFormatText());
		// LINEST
		assertEquals("2", Ranges.range(sheet, "B103").getCellFormatText());
		// LOGINV
		assertEquals("4.00", Ranges.range(sheet, "B106").getCellFormatText());
		// LOGNORMDIST
		assertEquals("0.04", Ranges.range(sheet, "B108").getCellFormatText());
		// MAX
		assertEquals("27", Ranges.range(sheet, "B110").getCellFormatText());
		// MAXA
		assertEquals("0.5", Ranges.range(sheet, "B112").getCellFormatText());
		// MEDIAN
		assertEquals("3", Ranges.range(sheet, "B114").getCellFormatText());
		// MIN
		assertEquals("2", Ranges.range(sheet, "B116").getCellFormatText());
		// MINA
		assertEquals("0", Ranges.range(sheet, "B118").getCellFormatText());
		// MODE
		assertEquals("4", Ranges.range(sheet, "B121").getCellFormatText());
		// NEGBINOMDIST
		assertEquals("0.06", Ranges.range(sheet, "B123").getCellFormatText());
		// NORMDIST
		assertEquals("0.91", Ranges.range(sheet, "B125").getCellFormatText());
		// MORMINV
		assertEquals("42.00", Ranges.range(sheet, "B127").getCellFormatText());
		// NORMSDIST
		assertEquals("0.91", Ranges.range(sheet, "B129").getCellFormatText());
		// NORMSINV
		assertEquals("1.33", Ranges.range(sheet, "B131").getCellFormatText());
		// PEARSON
		assertEquals("0.70", Ranges.range(sheet, "B133").getCellFormatText());
		// PERCENTILE
		assertEquals("1.9", Ranges.range(sheet, "B136").getCellFormatText());
		// PERCENTRANK
		assertEquals("0.33", Ranges.range(sheet, "B138").getCellFormatText());
		// PERMUT
		assertEquals("970200", Ranges.range(sheet, "B140").getCellFormatText());
		// POISSON
		assertEquals("0.12", Ranges.range(sheet, "B142").getCellFormatText());
		// PROB
		assertEquals("0.1", Ranges.range(sheet, "B144").getCellFormatText());
		// QUARTILE
		assertEquals("3.5", Ranges.range(sheet, "B147").getCellFormatText());
		// RANK
		assertEquals("3", Ranges.range(sheet, "B149").getCellFormatText());
		// RSQ
		assertEquals("0.061", Ranges.range(sheet, "B151").getCellFormatText());
		// SKEW
		assertEquals("0.36", Ranges.range(sheet, "B154").getCellFormatText());
		// SLOPE
		assertEquals("0.31", Ranges.range(sheet, "B156").getCellFormatText());
		// SMALL
		assertEquals("4", Ranges.range(sheet, "B159").getCellFormatText());
		// STANDARDIZE
		assertEquals("1.33", Ranges.range(sheet, "B162").getCellFormatText());
		// STDEV
		assertEquals("27.46", Ranges.range(sheet, "B164").getCellFormatText());
		// STDEVA
		assertEquals("27.46", Ranges.range(sheet, "B166").getCellFormatText());
		// STDEVP
		assertEquals("26.05", Ranges.range(sheet, "B168").getCellFormatText());
		// STDEVPA
		assertEquals("26.05", Ranges.range(sheet, "B170").getCellFormatText());
		// STEYX
		assertEquals("3.31", Ranges.range(sheet, "B172").getCellFormatText());
		// TDIST
		assertEquals("0.05", Ranges.range(sheet, "B175").getCellFormatText());
		// TINV
		assertEquals("1.96", Ranges.range(sheet, "B177").getCellFormatText());
		// TREND
		assertEquals("133953.33", Ranges.range(sheet, "B179").getCellFormatText());
		// TRIMMEAN
		assertEquals("3.78", Ranges.range(sheet, "B183").getCellFormatText());
		// TTEST
		assertEquals("0.20", Ranges.range(sheet, "B186").getCellFormatText());
		// VAR
		assertEquals("754.27", Ranges.range(sheet, "B189").getCellFormatText());
		// VARA
		assertEquals("754.27", Ranges.range(sheet, "B191").getCellFormatText());
		// VARP
		assertEquals("678.84", Ranges.range(sheet, "B193").getCellFormatText());
		// VARPA
		assertEquals("678.84", Ranges.range(sheet, "B195").getCellFormatText());
		// WEIBULL
		assertEquals("0.93", Ranges.range(sheet, "B197").getCellFormatText());
		// ZTEST
		assertEquals("0.09", Ranges.range(sheet, "B199").getCellFormatText());
	}
	
	@Ignore("ZSS-262")
	@Test
	public void testZSS262() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("Sheet1");
		Ranges.range(sheet, "A1").setCellEditText("1234.567");
		Ranges.range(sheet, "B1").setCellEditText("DOLLAR(A1,2)");
		assertEquals("NT$1,234.57", Ranges.range(sheet, "B1").getCellFormatText());
	}
	
	@Ignore("ZSS-263")
	@Test
	public void testZSS263() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("Sheet1");
		Ranges.range(sheet, "A1").setCellEditText("VALUE(\"16:48:00\")");
		assertEquals("0.7", Ranges.range(sheet, "A1").getCellFormatText());
	}
	
	@Ignore("ZSS-264")
	@Test
	public void testZSS264() throws IOException {
		final String filename = "book/264-text-formula.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("formula-text");
		assertEquals("EXCEL", Ranges.range(sheet, "B3").getCellFormatText());
		assertEquals("A", Ranges.range(sheet, "B5").getCellFormatText());
		assertEquals("text", Ranges.range(sheet, "B7").getCellFormatText());
		assertEquals("65", Ranges.range(sheet, "B10").getCellFormatText());
		assertEquals("ZK", Ranges.range(sheet, "B12").getCellFormatText());
		assertEquals("$1,234.57", Ranges.range(sheet, "B14").getCellFormatText());
		assertEquals("TRUE", Ranges.range(sheet, "B16").getCellFormatText());
		assertEquals("1", Ranges.range(sheet, "B20").getCellFormatText());
		assertEquals("1", Ranges.range(sheet, "B22").getCellFormatText());
		assertEquals("1,234.6", Ranges.range(sheet, "B24").getCellFormatText());
		assertEquals("\u6771 \u4EAC", Ranges.range(sheet, "B27").getCellFormatText());
		assertEquals("\u6771", Ranges.range(sheet, "B29").getCellFormatText());
		assertEquals("11", Ranges.range(sheet, "B31").getCellFormatText());
		assertEquals("11", Ranges.range(sheet, "B33").getCellFormatText());
		assertEquals("e. e. cummings", Ranges.range(sheet, "B35").getCellFormatText());
		assertEquals("Fluid", Ranges.range(sheet, "B38").getCellFormatText());
		assertEquals("Fluid", Ranges.range(sheet, "B40").getCellFormatText());
		assertEquals("This Is A Title", Ranges.range(sheet, "B42").getCellFormatText());
		assertEquals("abcde*k", Ranges.range(sheet, "B45").getCellFormatText());
		assertEquals("abcde*k", Ranges.range(sheet, "B47").getCellFormatText());
		assertEquals("*-*-*-", Ranges.range(sheet, "B49").getCellFormatText());
		assertEquals("Price", Ranges.range(sheet, "B51").getCellFormatText());
		assertEquals("Price", Ranges.range(sheet, "B53").getCellFormatText());
		assertEquals("6", Ranges.range(sheet, "B55").getCellFormatText());
		assertEquals("6", Ranges.range(sheet, "B57").getCellFormatText());
		assertEquals("Quarter 2, 2008", Ranges.range(sheet, "B59").getCellFormatText());
		assertEquals("Sale Price", Ranges.range(sheet, "B62").getCellFormatText());
		assertEquals("Date: 2007-08-06", Ranges.range(sheet, "B65").getCellFormatText());
		assertEquals("revenue in quarter 1", Ranges.range(sheet, "B68").getCellFormatText());
		assertEquals("TOTAL", Ranges.range(sheet, "B70").getCellFormatText());
		assertEquals("0.7", Ranges.range(sheet, "B73").getCellFormatText());
	}
	
	/**
	 * some "accounting" data format causes NullPointerException
	 */
	@Ignore("format text is wrong")
	@Test
	public void testZSS254() throws IOException {
		final String filename = "book/254-accounting.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("cell-data");
		assertEquals("NT$1,234.56", Ranges.range(sheet, "C2").getCellFormatText());
		
		// FIXME
		assertEquals("¥1,234.00", Ranges.range(sheet, "D2").getCellFormatText());
		// Excel display 06:12:00 PM in function field
		// ZSS display 0.75833333333 in function field
		assertEquals("06:12 PM", Ranges.range(sheet, "F2").getCellFormatText());
	}
	
	@Test
	public void testZSS334() throws IOException {
		// FIXME
		String bookName1 = "book/334-book1.xlsx";
		Book book1 = Importers.getImporter().imports(Issue200Test.class.getResourceAsStream(bookName1), bookName1);
		Sheet sheet1 = book1.getSheet("Sheet1");
		Ranges.range(sheet1, "A1").setCellValue("123");
		
		String bookName2 = "book/334-book2.xlsx";
		Book book2 = Importers.getImporter().imports(Issue200Test.class.getResourceAsStream(bookName2), bookName2);
		Sheet sheet2 = book2.getSheet("Sheet1");
		
		BookSeriesBuilder.getInstance().buildBookSeries(book1,book2);
		
		String bookName3 = "book/334-book3.xlsx";
		Book book3 = Importers.getImporter().imports(Issue200Test.class.getResourceAsStream(bookName3), bookName3);
		Sheet sheet3 = book2.getSheet("Sheet1");

	}
	
	/**
	 * When importing a xlsx file with images exported from Spreadsheet 2nd time, an exception is thrown
	 */
	@Test
	public void testZSS317() throws IOException {
		final String filename = "book/317-exportImage.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("export1");
		SheetOperationUtil.addPicture(Ranges.range(sheet), new AImage(new File(ShiftTest.class.getResource("").getPath() + "book/zklogo.png")));
		Util.export(workbook, Setup.getTempFile());
		Util.export(workbook, Setup.getTempFile());
	}
	
	/**
	 * Exception when delete columns
	 */
	@Test
	public void testZSS245() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("Sheet1");
		Ranges.range(sheet, "A1").toRowRange().delete(DeleteShift.LEFT);
	}
	
	/**
	 * When given start date is later than end date In NETWORKDAYS(), it returns #NAME?
	 */
	@Test
	public void testZSS355() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("Sheet1");
		Ranges.range(sheet, "A1").setCellEditText("=NETWORKDAYS(DATE(2013,6,2),DATE(2013,6,1))");
		assertEquals("#VALUE!", Ranges.range(sheet, "A1").getCellData().getFormatText());
	}
	
	/**
	 * AVERAGEIF, AVERAGEEIFS -> #NAME?
	 */
	@Test
	public void testZSS341() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("Sheet1");
		
		Ranges.range(sheet, "B11").setCellEditText("=AVERAGEIF(B12:E12,\"<23000\")");
		Ranges.range(sheet, "B12").setCellEditText("7000");
		Ranges.range(sheet, "C12").setCellEditText("14000");
		Ranges.range(sheet, "D12").setCellEditText("21000");
		Ranges.range(sheet, "E12").setCellEditText("28000");
		assertEquals("#NAME?", Ranges.range(sheet, "B11").getCellData().getFormatText());
		Ranges.range(sheet, "C11").setCellEditText("=AVERAGEIFS(B12:E12,B12:E12,\"<23000\",B12:E12,\">13000\")");
		assertEquals("#NAME?", Ranges.range(sheet, "C11").getCellData().getFormatText());
		
	}
	
	/**
	 * May need to check
	 * 1. HOUR, SECOND, MINUTE -> #VALUE!
	 * 2. TIMEVALUE -> #NAME?
	 * 3. YEARFRAC tolerance 0.005
	 */
	@Test
	public void testZSS267() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("Sheet1");
		
		// DATE
		Ranges.range(sheet, "B4").setCellEditText("2008");
		Ranges.range(sheet, "C4").setCellEditText("1");
		Ranges.range(sheet, "D4").setCellEditText("1");
		Ranges.range(sheet, "B3").setCellEditText("=DATE(B4,C4,D4)");
		CellOperationUtil.applyDataFormat(Ranges.range(sheet, "B3"), "Y/M/D");
		assertEquals("08/1/1", Ranges.range(sheet, "B3").getCellData().getFormatText());// '08/01/01' is correct and fixed in zpoi 3.9
		
		// DATEVALUE
		Ranges.range(sheet, "B6").setCellEditText("=DATEVALUE(\"2008/1/1\")");
		assertEquals("39448", Ranges.range(sheet, "B6").getCellData().getFormatText());
		
		// DAY
		Ranges.range(sheet, "B9").setCellEditText("2008/4/15");
		Ranges.range(sheet, "B8").setCellEditText("=DAY(B9)");
		assertEquals("15", Ranges.range(sheet, "B8").getCellData().getFormatText());
		
		// DAYS360
		Ranges.range(sheet, "B11").setCellEditText("2008/1/30");
		Ranges.range(sheet, "C11").setCellEditText("2008/2/1");
		Ranges.range(sheet, "B10").setCellEditText("=DAYS360(B11,C11)");
		assertEquals("1", Ranges.range(sheet, "B10").getCellData().getFormatText());
		
		// HOUR
		Ranges.range(sheet, "B13").setCellEditText("=HOUR(\"15:30\")");
		assertEquals("15", Ranges.range(sheet, "B13").getCellData().getFormatText());
		
		// MINUTE
		Ranges.range(sheet, "B15").setCellEditText("=MINUTE(\"4:48:00 PM\")");
		assertEquals("48", Ranges.range(sheet, "B15").getCellData().getFormatText());
		
		// MONTH
		Ranges.range(sheet, "B17").setCellEditText("=MONTH(DATE(2008,4,15))");
		assertEquals("4", Ranges.range(sheet, "B17").getCellData().getFormatText());
		
		// NETWORKDAYS
		Ranges.range(sheet, "B19").setCellEditText("=NETWORKDAYS(DATE(2013,4,1),DATE(2013,4,30))");
		assertEquals("22", Ranges.range(sheet, "B19").getCellData().getFormatText());
		
		// SECOND
		Ranges.range(sheet, "B23").setCellEditText("=SECOND(\"4:48 PM\")");
		assertEquals("0", Ranges.range(sheet, "B23").getCellData().getFormatText());
		
		// TIME
		Ranges.range(sheet, "B25").setCellEditText("=TIME(12,0,0)");
		assertEquals("0.5", Ranges.range(sheet, "B25").getCellData().getFormatText());
		
		// TIMEVALUE
		Ranges.range(sheet, "B27").setCellEditText("=TIMEVALUE(\"2:24 AM\")");
		assertEquals("#NAME?", Ranges.range(sheet, "B27").getCellData().getFormatText());
		
		// WEEKDAY
		Ranges.range(sheet, "B31").setCellEditText("=WEEKDAY(DATE(2008,2,14))");
		assertEquals("5", Ranges.range(sheet, "B31").getCellData().getFormatText());
		
		// WORKDAY
		Ranges.range(sheet, "B33").setCellEditText("=WORKDAY(DATE(2013,4,1),5)");
		assertEquals("41372", Ranges.range(sheet, "B33").getCellData().getFormatText());
		
		// YEAR
		Ranges.range(sheet, "B35").setCellEditText("=YEAR(DATE(2008,1,1))");
		assertEquals("2008", Ranges.range(sheet, "B35").getCellData().getFormatText());
		
		// YEARFRAC
		Ranges.range(sheet, "B37").setCellEditText("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))");
		assertEquals(0.58, Ranges.range(sheet, "B37").getCellData().getDoubleValue(), 0.005);
		
	}
	
	/**
	 * SERIESSUM(B105,0,2,B106:B109) is a unsupported function, cell value should be ERROR_NAME
	 */
	@Test
	public void testZSS261() throws IOException {
		final String filename = "book/math.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("formula-math");
		Range B104 = Ranges.range(sheet1, "B104");
		assertEquals("#NAME?", B104.getCellData().getFormatText());
	}
	
	@Ignore("ZSS-265")
	@Test
	public void testZSS265() throws IOException {
		final String filename = "book/266-info-formula.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("formula-info");
		assertEquals("3", Ranges.range(sheet, "B3").getCellFormatText());
		assertEquals("1", Ranges.range(sheet, "B5").getCellFormatText());
		assertEquals("12.0", Ranges.range(sheet, "B8").getCellFormatText());
		assertEquals("FALSE", Ranges.range(sheet, "B10").getCellFormatText());
		assertEquals("TRUE", Ranges.range(sheet, "B12").getCellFormatText());
		assertEquals("TRUE", Ranges.range(sheet, "B15").getCellFormatText());
		assertEquals("TRUE", Ranges.range(sheet, "B18").getCellFormatText());
		assertEquals("FALSE", Ranges.range(sheet, "B21").getCellFormatText());
		assertEquals("FALSE", Ranges.range(sheet, "B23").getCellFormatText());
		assertEquals("TRUE", Ranges.range(sheet, "B26").getCellFormatText());
		assertEquals("TRUE", Ranges.range(sheet, "B29").getCellFormatText());
		assertEquals("FALSE", Ranges.range(sheet, "B31").getCellFormatText());
		assertEquals("FALSE", Ranges.range(sheet, "B33").getCellFormatText());
		assertEquals("TRUE", Ranges.range(sheet, "B35").getCellFormatText());
		assertEquals("7", Ranges.range(sheet, "B38").getCellFormatText());
		assertEquals("#N/A", Ranges.range(sheet, "B41").getCellFormatText());
		assertEquals("2", Ranges.range(sheet, "B43").getCellFormatText());
	}
	
	/**
	 * ISERR
	 */
	@Test
	public void testZSS266() throws IOException {
		final String filename = "book/266-info-formula.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("formula-info");
		Range B12 = Ranges.range(sheet1, "B12");
		assertTrue(B12.getCellData().getBooleanValue());
	}
	
	/**
	 * test POI function without argument
	 */
	@Test
	public void testZSS342_1() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range A1 = Ranges.range(sheet1, "A1");
		A1.setCellEditText("=ACCRINT()");
		assertEquals("#VALUE!", A1.getCellData().getFormatText());
	}
	
	/**
	 * test ZK function without argument
	 */
	@Test
	public void testZSS342_2() throws IOException {
		final String filename = "book/blank.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		try {
			Range A1 = Ranges.range(sheet1, "A1");
			A1.setCellEditText("=ISERR()");
			assertEquals("#VALUE!", A1.getCellData().getFormatText());
		} catch(IllegalFormulaException e) {
			return;
		}
		fail(); // fail if no catch
	}
	
	/**
	 * if a sheet contains "data validation" configuration with empty criteria, 
	 * it causes null pointer exception when displaying it, then crashed.
	 */
	@Test
	public void testZSS255() throws IOException {
		final String filename = "book/255-cell-data.xlsx";
		Util.loadBook(this,filename);
	}
	
	/**
	 * validation cell
	 */
	@Test
	public void testZSS260() throws IOException {
		final String filename = "book/260-validation.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("Validation");
		Range C3 = Ranges.range(sheet1, "C3");
		C3.setCellEditText("2");
		
		Range C5 = Ranges.range(sheet1, "C5");
		C5.setCellEditText("2013/1/2");
	}
	
	/**
	 * Range.getRow() return unexpected row number
	 */
	@Test
	public void testZSS275() {
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range rangeA = Ranges.range(sheet1, "H11:J13");
		assertEquals(10, rangeA.getRow());
	}
	
	/**
	 * 1. merge a 3 x 3 rangeA
	 * 2. create a whole row rangeB which cross the merged cell
	 * 3. perform unmerge on the rangeB
	 * 4. rangeA should be unmerged
	 */
	@Test
	public void testZSS395_1() {
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range rangeA = Ranges.range(sheet1, "H11:J13");
		rangeA.merge(false); // merge a 3 x 3
		AssertUtil.assertMergedRange(rangeA); // it should be merged
		Range rangeB = Ranges.range(sheet1, "H12"); // a whole row cross the merged cell
		rangeB.toRowRange().unmerge(); // perform unmerge operation
		AssertUtil.assertNotMergedRange(rangeA); // should be unmerged now
	}
	
	/**
	 * 1. merge a 3 x 3 rangeA
	 * 2. create a whole column rangeB which cross the merged cell
	 * 3. perform unmerge on the rangeB
	 * 4. rangeA should be unmerged
	 */
	@Test
	public void testZSS395_2() {
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range rangeA = Ranges.range(sheet1, "H11:J13");
		rangeA.merge(false); // merge a 3 x 3
		AssertUtil.assertMergedRange(rangeA); // it should be merged
		Range rangeB = Ranges.range(sheet1, "H12"); // a whole row cross the merged cell
		rangeB.toColumnRange().unmerge(); // perform unmerge operation
		AssertUtil.assertNotMergedRange(rangeA); // should be unmerged now
	}
	
	/**
	 * cut a merged cell and paste to another cell, the original cell doesn't become unmerged cells
	 *     1
	 *  2  5  7
	 *  4     9
	 * 1 is a horizontal merged cell 1 x 3.
	 * 5 is a vertical merged cell 2 x 1.
	 * source (H11,J13) to destination (C5, E7)
	 */
	@Test
	public void testZSS301() {
		
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		
		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_BOTTOM_ROW = 12;
		int SRC_RIGHT_COL = 9;
		int SRC_ROW_COUNT = SRC_BOTTOM_ROW - SRC_TOP_ROW + 1;
		int SRC_COL_COUNT = SRC_RIGHT_COL - SRC_LEFT_COL + 1;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H11, J11), horizontal merge, cell value is 1
		Range range_H11J11 = Ranges.range(sheet1, 10, 7, 10, 9);
		range_H11J11.merge(false);
		
		// Merge (I12, I13), vertical merge, cell value is 5
		Range range_I12I13 = Ranges.range(sheet1, 11, 8, 12, 8);
		range_I12I13.merge(false);
		
		// should be merged region
		AssertUtil.assertMergedRange(range_H11J11);
		AssertUtil.assertMergedRange(range_I12I13);
		
		int dstTopRow = 4;
		int dstLeftCol = 2;
		int dstBottomRow = 6;
		int dstRightCol = 4;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
		// validate destination
		
		// this should be a merged cell
		Range horizontalMergedRange = Ranges.range(sheet1, 4, 2, 4, 4);
		assertEquals(1, horizontalMergedRange.getCellData().getDoubleValue(), 1E-8);
		AssertUtil.assertMergedRange(horizontalMergedRange);
		
		// this should be a merged cell
		Range verticalMergedRange = Ranges.range(sheet1, 5, 3, 6, 3);
		assertEquals(5, verticalMergedRange.getCellData().getDoubleValue(), 1E-8);
		AssertUtil.assertMergedRange(verticalMergedRange);
		
		assertEquals(4, Ranges.range(sheet1, 5, 2).getCellData().getDoubleValue(), 1E-8);
		assertEquals(7, Ranges.range(sheet1, 6, 2).getCellData().getDoubleValue(), 1E-8);
		assertEquals(6, Ranges.range(sheet1, 5, 4).getCellData().getDoubleValue(), 1E-8);
		assertEquals(9, Ranges.range(sheet1, 6, 4).getCellData().getDoubleValue(), 1E-8);
		
		// validate source
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL+2).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+1, SRC_LEFT_COL).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+1, SRC_LEFT_COL+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+1, SRC_LEFT_COL+2).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+2, SRC_LEFT_COL).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+2, SRC_LEFT_COL+1).getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, SRC_TOP_ROW+2, SRC_LEFT_COL+2).getCellData().getType().ordinal(), 1E-8);

		// should not be merged region anymore
		AssertUtil.assertNotMergedRange(range_H11J11);
		AssertUtil.assertNotMergedRange(range_I12I13);
	}
	
	/**
	 * Delete multiple columns causes a exception
	 * Delete column E F
	 */
	@Test
	public void testZSS303() {
		
		final String filename = "book/shiftTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_E = Ranges.range(sheet1, "E3:F3");
		range_E.toColumnRange().delete(DeleteShift.DEFAULT);
			
		assertEquals("G3", Ranges.range(sheet1, "E3").getCellData().getEditText());
		assertEquals("G4", Ranges.range(sheet1, "E4").getCellData().getEditText());
		assertEquals("G5", Ranges.range(sheet1, "E5").getCellData().getEditText());
		assertEquals("G8", Ranges.range(sheet1, "E8").getCellData().getEditText());
		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F5").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "F8").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G4").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G5").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G8").getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * perform "clear style" on a merged cell doesn't split it back to original unmerged cells
	 */
	@Test
	public void testZSS298() {
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("Sheet1");
		Range range = Ranges.range(sheet, "H11:J13");
		range.merge(false); // 1. merge
		AssertUtil.assertMergedRange(range); // is it merged?
		CellOperationUtil.clearStyles(range);
		AssertUtil.assertNotMergedRange(range); // is it unmerged?
	}
	
	/**
	 * Insert a chart after deleting first chart causes an exception
	 */
	@Test
	public void testZSS326() throws IOException {
		final String filename = "book/insert-charts.xlsx";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("chart-image");
		// Insert chart 1, 2
		ChartData cd1 = ChartDataUtil.getChartData(sheet, new AreaRef(4,1,14,1), Chart.Type.LINE);
		Chart chart1 = SheetOperationUtil.addChart(Ranges.range(sheet, "A1"), cd1, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		ChartData cd2 = ChartDataUtil.getChartData(sheet, new AreaRef(4,2,14,2), Chart.Type.LINE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "I1"), cd2, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		
		// Delete chart 1
		SheetOperationUtil.deleteChart(Ranges.range(sheet, "A1"), chart1);
		
		// Insert chart 3
		ChartData cd3 = ChartDataUtil.getChartData(sheet, new AreaRef(4,3,14,3), Chart.Type.LINE);
		SheetOperationUtil.addChart(Ranges.range(sheet, "Q1"), cd3, Chart.Type.LINE, Grouping.STANDARD, LegendPosition.TOP);
		
	}
	
	/**
	 * after merging and unmerging cells, input and style both work incorrect
	 */
	@Test
	public void testZSS290() {
		final String filename = "book/290-merge.xls";
		Book workbook = Util.loadBook(this,filename);
		Sheet sheet = workbook.getSheet("cell");
		Range range = Ranges.range(sheet, "B5:D7");
		range.merge(true);
		range.merge(false);
		range.unmerge();
		String color = "#548dd4";
		assertEquals(color, Ranges.range(sheet, "B5").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "B6").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "B7").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "C5").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "C6").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "C7").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "D5").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "D6").getCellStyle().getBackgroundColor().getHtmlColor());
		assertEquals(color, Ranges.range(sheet, "D7").getCellStyle().getBackgroundColor().getHtmlColor());
	}
	
	/**
	 *     1
	 *  2  5  7
	 *  4     9
	 * 1 is a horizontal merged cell 1 x 3.
	 * 5 is a vertical merged cell 2 x 1.
	 * transpose paste.
	 * source (H11,J13) to destination (C5, E7).
	 */
	@Test 
	public void testZSS277() {
		
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		
		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_BOTTOM_ROW = 12;
		int SRC_RIGHT_COL = 9;
		int SRC_ROW_COUNT = SRC_BOTTOM_ROW - SRC_TOP_ROW + 1;
		int SRC_COL_COUNT = SRC_RIGHT_COL - SRC_LEFT_COL + 1;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H11, J11), horizontal merge, cell value is 1
		Range range_H11J11 = Ranges.range(sheet1, 10, 7, 10, 9);
		range_H11J11.merge(false);
		
		// Merge (I12, I13), vertical merge, cell value is 5
		Range range_I12I13 = Ranges.range(sheet1, 11, 8, 12, 8);
		range_I12I13.merge(false);
		
		// should be merged region
		AssertUtil.assertMergedRange(range_H11J11);
		AssertUtil.assertMergedRange(range_I12I13);
		
		int dstTopRow = 4;
		int dstLeftCol = 2;
		int dstBottomRow = 6;
		int dstRightCol = 4;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.pasteSpecial(dstRange, PasteType.ALL, PasteOperation.NONE, false, true);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);
		
		// validate destination
		
		// this should be a merged cell
		// origin is horizontal, but now vertical
		Range horizontalMergedRange = Ranges.range(sheet1, 4, 2, 6, 2);
		
		assertEquals(1, horizontalMergedRange.getCellData().getDoubleValue(), 1E-8);
		AssertUtil.assertMergedRange(horizontalMergedRange);
		
		// this should be a merged cell
		// origin is vertical, but now horizontal
		Range verticalMergedRange = Ranges.range(sheet1, 5, 3, 5, 4);
		assertEquals(5, verticalMergedRange.getCellData().getDoubleValue(), 1E-8);
		AssertUtil.assertMergedRange(verticalMergedRange);
		
		assertEquals(4, Ranges.range(sheet1, dstTopRow, dstLeftCol+1).getCellData().getDoubleValue(), 1E-8);
		assertEquals(7, Ranges.range(sheet1, dstTopRow, dstLeftCol+2).getCellData().getDoubleValue(), 1E-8);

		assertEquals(6, Ranges.range(sheet1, dstTopRow+2, dstLeftCol+1).getCellData().getDoubleValue(), 1E-8);
		assertEquals(9, Ranges.range(sheet1, dstTopRow+2, dstLeftCol+2).getCellData().getDoubleValue(), 1E-8);		
	}
	
	/**
	 * paste repeat overlap, left cover
	 * source (H11, J13) 3 x 3
	 * dst (H7, J12) 6 x 3
	 */
	@Test 
	public void testPasteZSS300() {
		
		
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);

		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_BOTTOM_ROW = 12;
		int SRC_RIGHT_COL = 9;
		int SRC_ROW_COUNT = SRC_BOTTOM_ROW - SRC_TOP_ROW + 1;
		int SRC_COL_COUNT = SRC_RIGHT_COL - SRC_LEFT_COL + 1;
		int dstTopRow = 6;
		int dstLeftCol = 7;
		int dstBottomRow = 11;
		int dstRightCol = 9;
		
		// operation
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_BOTTOM_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol, dstBottomRow, dstRightCol);
		Range pasteRange = srcRange.paste(dstRange); // copy src to dst
		
		int realDstRowCount = pasteRange.getRowCount();
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(2, realDstRowCount / SRC_ROW_COUNT);
		assertEquals(1, realDstColCount / SRC_COL_COUNT);			
		
		validateZSS300(sheet1, dstTopRow, dstLeftCol, realDstRowCount / SRC_ROW_COUNT, realDstColCount / SRC_COL_COUNT);
	}
	
	/**
	 * E3:E5 shift up
	 */
	@Test
	public void testZSS389_1() {

		final String filename = "book/shiftTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_E3E5 = Ranges.range(sheet1, "E3:E5");
		range_E3E5.delete(DeleteShift.UP);
		
		// validate E1, E2 is still correct
		assertEquals("E1", Ranges.range(sheet1, "E1").getCellData().getEditText());
		assertEquals("E2", Ranges.range(sheet1, "E2").getCellData().getEditText());
		
		// validate shift up
		assertEquals("E6", Ranges.range(sheet1, "E3").getCellData().getEditText());
		assertEquals("E7", Ranges.range(sheet1, "E4").getCellData().getEditText());
		assertEquals("E8", Ranges.range(sheet1, "E5").getCellData().getEditText());
		
		// validate E6, E7, E8 is blank
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "E8").getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * G3:G5 shift up
	 */
	@Test
	public void testZSS389_2() {
		
		final String filename = "book/shiftTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		Range range_G3G5 = Ranges.range(sheet1, "G3:G5");
		range_G3G5.delete(DeleteShift.UP);
		
		// validate shift up
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G3").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G4").getCellData().getType().ordinal(), 1E-8);
		assertEquals("G8", Ranges.range(sheet1, "G5").getCellData().getEditText());
		
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G6").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G7").getCellData().getType().ordinal(), 1E-8);
		assertEquals(CellType.BLANK.ordinal(), Ranges.range(sheet1, "G8").getCellData().getType().ordinal(), 1E-8);
	}
	
	/**
	 * 1 2 3
	 *   4
	 * 5 6 7
	 * 
	 * 4 is a horizontal merged cell 1 x 3.
	 * paste 1 to 4.
	 * source (H11) to destination (H12).
	 * a single cell paste to a merged cell.
	 */
	@Test
	public void testZSS315_1() {
		
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);
		
		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H12, J12), horizontal merge, cell value is 4
		Range range_H12J12 = Ranges.range(sheet1, 11, 7, 11, 9);
		range_H12J12.merge(false);
		
		assertEquals(4, range_H12J12.getCellData().getDoubleValue(), 1E-8);
		
		// should be merged region
		AssertUtil.assertMergedRange(range_H12J12);
		
		int dstTopRow = 11;
		int dstLeftCol = 7;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / 1);
		assertEquals(1, realDstColCount / 1);
		
		// this should still be a merged cell
		AssertUtil.assertMergedRange(range_H12J12);
		
		// value validation
		assertEquals(1, pasteRange.getCellData().getDoubleValue(), 1E-8);		
	}
	
	/**
	 *   1
	 * 4 5 6
	 * 7 8 9
	 * 1 is a horizontal merged cell 1 x 3.
	 * paste 1 to 4.
	 * source (H11) to destination (H12).
	 * a single merged cell paste to a single cell.
	 */
	@Test
	public void testZSS315_2() {
		
		final String filename = "book/pasteTest.xlsx";
		Book workbook = Util.loadBook(this,filename);

		int SRC_TOP_ROW = 10;
		int SRC_LEFT_COL = 7;
		int SRC_RIGHT_COL = 9;
		
		Sheet sheet1 = workbook.getSheet("Sheet1");
		
		// preparation
		// Merge (H11, J11), horizontal merge, cell value is 1
		Range range_H11J11 = Ranges.range(sheet1, 10, 7, 10, 9);
		range_H11J11.merge(false);
		
		assertEquals(1, range_H11J11.getCellData().getDoubleValue(), 1E-8);
		
		// should be merged region
		AssertUtil.assertMergedRange(range_H11J11);
		
		int dstTopRow = 11;
		int dstLeftCol = 7;
		
		Range srcRange = Ranges.range(sheet1, SRC_TOP_ROW, SRC_LEFT_COL, SRC_TOP_ROW, SRC_RIGHT_COL);
		Range dstRange = Ranges.range(sheet1, dstTopRow, dstLeftCol);
		Range pasteRange = srcRange.paste(dstRange);
		
		int realDstRowCount = pasteRange.getRowCount();;
		int realDstColCount = pasteRange.getColumnCount();
		
		// paste size validation
		assertEquals(1, realDstRowCount / 1);
		assertEquals(1, realDstColCount / 3);
		
		// this should be a merged cell
		AssertUtil.assertMergedRange(pasteRange);
		
		// value validation
		assertEquals(1, pasteRange.getCellData().getDoubleValue(), 1E-8);
		
	}
	
	private void validateZSS300(Sheet sheet1, int dstTopRow, int dstLeftCol, int rowRepeat, int colRepeat) {
		// validation
		for(int rr = rowRepeat, row = dstTopRow; rr > 0; rr--) {
			for(int cc = colRepeat, col = dstLeftCol; cc > 0; cc--) {
				assertEquals(1, Ranges.range(sheet1, row, col).getCellData().getDoubleValue(), 1E-8);
				assertEquals(2, Ranges.range(sheet1, row, col+1).getCellData().getDoubleValue(), 1E-8);
				assertEquals(3, Ranges.range(sheet1, row, col+2).getCellData().getDoubleValue(), 1E-8);
				// next row
				assertEquals(4, Ranges.range(sheet1, row+1, col).getCellData().getDoubleValue(), 1E-8);
				assertEquals(5, Ranges.range(sheet1, row+1, col+1).getCellData().getDoubleValue(), 1E-8);
				assertEquals(6, Ranges.range(sheet1, row+1, col+2).getCellData().getDoubleValue(), 1E-8);
				// next row
				assertEquals(7, Ranges.range(sheet1, row+2, col).getCellData().getDoubleValue(), 1E-8);
				assertEquals(8, Ranges.range(sheet1, row+2, col+1).getCellData().getDoubleValue(), 1E-8);
				assertEquals(9, Ranges.range(sheet1, row+2, col+2).getCellData().getDoubleValue(), 1E-8);
				// move column pointer whenever column repeat
				col += 3;
			}
			// move row pointer whenever row repeat
			row += 3;
		}
	}
	
	
	//not fix yet
	@Test
	public void testZSS179() throws IOException {

		final String filename = "book/179-insertexception-simple.xlsx";//should also test non-simple one
		Book book = Util.loadBook(this, filename);

		
		Range r = Ranges.range(book.getSheetAt(0), "A1");
    	Assert.assertEquals("A", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "B1");
    	Assert.assertEquals("", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "C1");
    	Assert.assertEquals("B", r.getCellEditText());
		
		r = Ranges.range(book.getSheetAt(0), "B1");
		r.insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
		java.io.File temp = java.io.File.createTempFile("test",".xlsx");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	//export first time
    	exporter.export(book, fos);
    	r = Ranges.range(book.getSheetAt(0), "A1");
    	Assert.assertEquals("A", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "B1");
    	Assert.assertEquals("", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "C1");
    	Assert.assertEquals("", r.getCellEditText());
    	r = Ranges.range(book.getSheetAt(0), "D1");
    	Assert.assertEquals("B", r.getCellEditText());
		
	}

	@Test
	public void testZSS399() throws Exception {
		// load book
		final String filename = "book/399-pdf-gridline.xlsx";
		Book book = Util.loadBook(this, filename);
		

		// print setting >> with grid lines 
		Sheet sheet = book.getSheetAt(0);
		sheet.getPoiSheet().setPrintGridlines(true);
		
		// export to PDF
		File temp = Setup.getTempFile("zss399-1", ".pdf");
		FileOutputStream fos = new FileOutputStream(temp);
		Exporter pdfExporter = Exporters.getExporter("pdf");
		pdfExporter.export(book, fos);
		fos.close();
		System.out.println("xlsx with grid lines export to " + temp);
		
		// open PDF
		// Check list:
		// 1. correct content
		// 2. must have grid lines 
		Util.open(temp);
		
		// print setting >> without grid lines 
		sheet = book.getSheetAt(0);
		sheet.getPoiSheet().setPrintGridlines(false);
		
		// export to PDF
		temp = Setup.getTempFile("zss399-2", ".pdf");
		fos = new FileOutputStream(temp);
		pdfExporter = Exporters.getExporter("pdf");
		pdfExporter.export(book, fos);
		fos.close();
		System.out.println("xlsx without grid lines export to " + temp);

		// open PDF
		// Check list:
		// 1. correct content
		// 2. NO grid lines 
		Util.open(temp);
	}
	
	@Test
	public void testZSS356() throws Exception {
		Book book = Util.loadBook(this,"book/356-enable-autofilter.xlsx");
		//shouldn't throw exception
		Ranges.range(book.getSheetAt(0),"D18").enableAutoFilter(true);
	}
	
	@Test
	public void testZSS276() throws Exception {
		Book book = Util.loadBook(this,"book/276-format.xls");
		//shouldn't throw exception
		Range r = Ranges.range(book.getSheetAt(0),"E2");
		
		Assert.assertEquals("2013/4/12", r.getCellFormatText());
		
	}
	
	@Test
	public void testZSS214() throws Exception {
		Book book = Util.loadBook(this,"book/214-npe-rename.xlsx");
		//shouldn't throw exception
		Range r = Ranges.range(book.getSheet("sheet1"),"A1");
		
		Assert.assertEquals("Dennis", r.getCellFormatText());
		
		//shound't throw exception
		r.setSheetName("sheet1x");
		
		File temp = Setup.getTempFile();
		//export first time
		Exporters.getExporter().export(book, temp);
		
		book = Importers.getImporter().imports(temp, "test");
		
		Assert.assertNull(book.getSheet("sheet1"));
		
		r = Ranges.range(book.getSheet("sheet1x"),"A1");
		Assert.assertEquals("Dennis", r.getCellFormatText());
	}
	
	@Test
	public void testZSS323() throws Exception {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		//shouldn't throw exception
		Range r = Ranges.range(book.getSheetAt(0),"A1:D5");
		
		//shoun't throw exception
		r.sort(true);
		r.sort(false);
		
		
		book = Util.loadBook(this,"book/blank.xls");
		//shouldn't throw exception
		r = Ranges.range(book.getSheetAt(0),"A1:D5");
		
		//shoun't throw exception
		r.sort(true);
		r.sort(false);
	}
	
	@Test
	public void testZSS256() throws Exception {
		testZSS256(Util.loadBook(this,"book/256-rowcolumn.xlsx"));
		testZSS256(Util.loadBook(this,"book/256-rowcolumn.xls"));
	}
	
	
	private void testZSS256(Book book) throws Exception {
	
		//shouldn't throw exception
		Sheet sheet = book.getSheet("rowcolumn");

		Assert.assertFalse(sheet.isRowHidden(4));
		Assert.assertTrue(sheet.isRowHidden(5));
		Assert.assertFalse(sheet.isRowHidden(6));
		
		Assert.assertEquals(10, BookHelper.getMaxConfiguredColumn((XSheet)sheet.getPoiSheet()));
		Assert.assertFalse(sheet.isColumnHidden(3));
		Assert.assertTrue(sheet.isColumnHidden(4));
		Assert.assertFalse(sheet.isColumnHidden(5));
	}
}
