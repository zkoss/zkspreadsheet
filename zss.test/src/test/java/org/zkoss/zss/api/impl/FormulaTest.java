package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * TODO the test case Not included categories:
 * Math, Financial, Lookup, Logical
 * @author kuro, Hawk
 *
 */
public class FormulaTest {

	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}

	@Before
	public void startUp() throws Exception {
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN, -1));
	}

	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}

	
	// Date & Time --------------------------------------------------------------------------------
	@Test
	public void testDATE() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B4").setCellEditText("2008");
		Ranges.range(sheet, "C4").setCellEditText("1");
		Ranges.range(sheet, "D4").setCellEditText("1");
		Ranges.range(sheet, "B3").setCellEditText("=DATE(B4,C4,D4)");
		CellOperationUtil.applyDataFormat(Ranges.range(sheet, "B3"), "Y/M/D");
		assertEquals("08/1/1", Ranges.range(sheet, "B3").getCellData().getFormatText());// '08/01/01'// is correct and fixed in zpoi 3.9
	}

	@Test
	public void testDATEVALUE() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B6").setCellEditText("=DATEVALUE(\"2008/1/1\")");
		assertEquals("39448", Ranges.range(sheet, "B6").getCellData().getFormatText());
	}

	@Test
	public void testDAY() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B9").setCellEditText("2008/4/15");
		Ranges.range(sheet, "B8").setCellEditText("=DAY(B9)");
		assertEquals("15", Ranges.range(sheet, "B8").getCellData().getFormatText());
	}

	@Test
	public void testDAY360() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B11").setCellEditText("2008/1/30");
		Ranges.range(sheet, "C11").setCellEditText("2008/2/1");
		Ranges.range(sheet, "B10").setCellEditText("=DAYS360(B11,C11)");
		assertEquals("1", Ranges.range(sheet, "B10").getCellData().getFormatText());
	}

	@Test
	public void testHOUR() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "C13").setCellEditText("03:30:30");
		//must accept a cell reference to a time instead of a time string
		Ranges.range(sheet, "B13").setCellEditText("=HOUR(C13)");
		assertEquals("3", Ranges.range(sheet, "B13").getCellData().getFormatText());
	}

	@Test
	public void testMINUTE() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "C15").setCellEditText("4:48:00 PM");
		//must accept a cell reference to a time instead of a time string
		Ranges.range(sheet, "B15").setCellEditText("=MINUTE(C15)");
		assertEquals("48", Ranges.range(sheet, "B15").getCellData().getFormatText());
	}

	@Test
	public void testMONTH() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B17").setCellEditText("=MONTH(DATE(2008,4,15))");
		assertEquals("4", Ranges.range(sheet, "B17").getCellData().getFormatText());
	}

	@Test
	public void testNETWORKDAYS() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B19").setCellEditText("=NETWORKDAYS(DATE(2013,4,1),DATE(2013,4,30))");
		assertEquals("22", Ranges.range(sheet, "B19").getCellData().getFormatText());
	}

	@Test
	public void testSECOND() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B23").setCellEditText("=SECOND(\"4:48 PM\")");
		assertEquals("#VALUE!", Ranges.range(sheet, "B23").getCellData().getFormatText());
	}

	@Test
	public void testTIME() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B25").setCellEditText("=TIME(12,0,0)");
		assertEquals("0.5", Ranges.range(sheet, "B25").getCellData().getFormatText());
	}

	@Test
	public void testWEEKDAY() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B31").setCellEditText("=WEEKDAY(DATE(2008,2,14))");
		assertEquals("5", Ranges.range(sheet, "B31").getCellData().getFormatText());
	}

	@Test
	public void testWORKDAY() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B33").setCellEditText("=WORKDAY(DATE(2013,4,1),5)");
		assertEquals("41372", Ranges.range(sheet, "B33").getCellData().getFormatText());
	}

	@Test
	public void testYEAR() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B35").setCellEditText("=YEAR(DATE(2008,1,1))");
		assertEquals("2008", Ranges.range(sheet, "B35").getCellData().getFormatText());
	}

	@Test
	public void testYEARFRAC() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		Sheet sheet = book.getSheet("Sheet1");
		Ranges.range(sheet, "B37").setCellEditText("=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))");
		assertEquals(0.58, Ranges.range(sheet, "B37").getCellData().getDoubleValue(), 0.005);
	}

	
	// Text ---------------------------------------------------------------------------------------
	@Test
	public void testCHAR() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("A", Ranges.range(sheet, "B5").getCellFormatText());
	}

	@Test
	public void testCLEAN() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("text", Ranges.range(sheet, "B7").getCellFormatText());
	}

	@Test
	public void testLOWER() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("e. e. cummings", Ranges.range(sheet, "B35").getCellFormatText());
	}

	@Test
	public void testCODE() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("65", Ranges.range(sheet, "B10").getCellFormatText());
	}

	@Test
	public void testCONCATENATE() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("ZK", Ranges.range(sheet, "B12").getCellFormatText());
	}

	@Test
	public void testEXACT() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("TRUE", Ranges.range(sheet, "B16").getCellFormatText());
	}

	@Test
	public void testFIND() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B20").getCellFormatText());
	}

	@Test
	public void testFIXED() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1,234.6", Ranges.range(sheet, "B24").getCellFormatText());
	}

	@Test
	public void testLEFT() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("\u6771 \u4EAC ", Ranges.range(sheet, "B27").getCellFormatText());
	}

	@Test
	public void testLEN() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("11", Ranges.range(sheet, "B31").getCellFormatText());
	}

	@Test
	public void testMID() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Fluid", Ranges.range(sheet, "B38").getCellFormatText());
	}

	@Test
	public void testPROPER() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("This Is A Title", Ranges.range(sheet, "B42").getCellFormatText());
	}

	@Test
	public void testREPLACE() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("abcde*k", Ranges.range(sheet, "B45").getCellFormatText());
	}

	@Test
	public void testREPT() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("*-*-*-", Ranges.range(sheet, "B49").getCellFormatText());
	}

	@Test
	public void testRIGHT() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Price", Ranges.range(sheet, "B51").getCellFormatText());
	}

	@Test
	public void testSEARCH() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("6", Ranges.range(sheet, "B55").getCellFormatText());
	}

	@Test
	public void testSUBSTITUTE() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Quarter 2, 2008", Ranges.range(sheet, "B59").getCellFormatText());
	}

	@Test
	public void testT() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Sale Price", Ranges.range(sheet, "B62").getCellFormatText());
	}

	@Test
	public void testTEXT() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Date: 2007-08-06", Ranges.range(sheet, "B65").getCellFormatText());
	}

	@Test
	public void testTRIM() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("revenue in quarter 1", Ranges.range(sheet, "B68").getCellFormatText());
	}

	@Test
	public void testUPPER() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("TOTAL", Ranges.range(sheet, "B70").getCellFormatText());
	}

	@Test
	public void testVALUE() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("0.7", Ranges.range(sheet, "B73").getCellFormatText());
	}

	//Info ----------------------------------------------------------------------------------------
	@Test
	public void testERRORTYPE() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("1", Ranges.range(sheet, "B5").getCellFormatText());
	}

	@Test
	public void testISBLANK() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B10").getCellFormatText());
	}

	@Test
	public void testISERR() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B12").getCellFormatText());
	}

	@Test
	public void testISERROR() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B15").getCellFormatText());
	}

	@Test
	public void testISEVEN() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B18").getCellFormatText());
	}

	@Test
	public void testISLOGICAL() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B21").getCellFormatText());
	}

	@Test
	public void testISNA() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B23").getCellFormatText());
	}

	@Test
	public void testISNONTEXT() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B26").getCellFormatText());
	}

	@Test
	public void testISNUMBER() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B29").getCellFormatText());
	}

	@Test
	public void testISODD() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B31").getCellFormatText());
	}

	@Test
	public void testISREF() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B33").getCellFormatText());
	}

	@Test
	public void testISTEXT() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B35").getCellFormatText());
	}

	@Test
	public void testN() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("7", Ranges.range(sheet, "B38").getCellFormatText());
	}

	@Test
	public void testNA() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("#N/A", Ranges.range(sheet, "B41").getCellFormatText());
	}

	@Test
	public void testTYPE() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("2", Ranges.range(sheet, "B43").getCellFormatText());
	}

	
	
	//Statistical  --------------------------------------------------------------------------------

	@Test
	public void testAVEDEV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.02", Ranges.range(sheet, "B3").getCellFormatText());
	}

	@Test
	public void testAVERAGE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("11", Ranges.range(sheet, "B6").getCellFormatText());
	}
	
	@Test
	public void testAVERAGEA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.6", Ranges.range(sheet, "B8").getCellFormatText());
	}

	@Test
	public void testBIOMDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.21", Ranges.range(sheet, "B21").getCellFormatText());
	}

	@Test
	public void testCHIDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.05", Ranges.range(sheet, "B23").getCellFormatText());
	}

	@Test
	public void testCHIINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("18.31", Ranges.range(sheet, "B25").getCellFormatText());
	}

	@Test
	public void testCOUNT() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B41").getCellFormatText());
	}

	@Test
	public void testCOUNTA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("6", Ranges.range(sheet, "B43").getCellFormatText());
	}

	@Test
	public void testCOUNTBLANK() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B45").getCellFormatText());
	}

	@Test
	public void testCOUNTIF() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B47").getCellFormatText());
	}

	@Test
	public void testDEVSQ() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("48", Ranges.range(sheet, "B55").getCellFormatText());
	}

	@Test
	public void testEXPONDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.86", Ranges.range(sheet, "B57").getCellFormatText());
	}

	@Test
	public void testGAMMAINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("10.00", Ranges.range(sheet, "B78").getCellFormatText());
	}

	@Test
	public void testGAMMALN() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.79", Ranges.range(sheet, "B80").getCellFormatText());
	}

	@Test
	public void testGEOMEAN() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.48", Ranges.range(sheet, "B82").getCellFormatText());
	}

	@Test
	public void testHYPGEOMDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.36", Ranges.range(sheet, "B92").getCellFormatText());
	}

	@Test
	public void testKURT() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("-0.15", Ranges.range(sheet, "B97").getCellFormatText());
	}

	@Test
	public void testLARGE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5", Ranges.range(sheet, "B100").getCellFormatText());
	}

	@Test
	public void testMAX() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("27", Ranges.range(sheet, "B110").getCellFormatText());
	}

	@Test
	public void testMAXA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.5", Ranges.range(sheet, "B112").getCellFormatText());
	}

	@Test
	public void testMEDIAN() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B114").getCellFormatText());
	}

	@Test
	public void testMIN() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B116").getCellFormatText());
	}

	@Test
	public void testMINA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0", Ranges.range(sheet, "B118").getCellFormatText());
	}

	@Test
	public void testMODE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B121").getCellFormatText());
	}

	@Test
	public void testNORMDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.91", Ranges.range(sheet, "B125").getCellFormatText());
	}

	@Test
	public void testPOISSON() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.12", Ranges.range(sheet, "B142").getCellFormatText());
	}

	@Test
	public void testRANK() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B149").getCellFormatText());
	}

	@Test
	public void testSKEW() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.36", Ranges.range(sheet, "B154").getCellFormatText());
	}

	@Test
	public void testSLOPE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.31", Ranges.range(sheet, "B156").getCellFormatText());
	}

	@Test
	public void testSMALL() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B159").getCellFormatText());
	}

	@Test
	public void testSTDEV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("27.46", Ranges.range(sheet, "B164").getCellFormatText());
	}

	@Test
	public void testVAR() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("754.27", Ranges.range(sheet, "B189").getCellFormatText());
	}

	@Test
	public void testVARP() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("678.84", Ranges.range(sheet, "B193").getCellFormatText());
	}

	@Test
	public void testWEIBULL() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.93", Ranges.range(sheet, "B197").getCellFormatText());
	}
	
	
	
	//Engineering ---------------------------------------------------------------------------------
	@Test
	public void testBESSELI() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.98", Ranges.range(sheet, "B3").getCellFormatText());
	}

	@Test
	public void testBESSELJ() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.33", Ranges.range(sheet, "B5").getCellFormatText());
	}

	@Test
	public void testBESSELK() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.28", Ranges.range(sheet, "B7").getCellFormatText());
	}

	@Test
	public void testBESSELY() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.15", Ranges.range(sheet, "B9").getCellFormatText());
	}

	@Test
	public void testBIN2DEC() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("100", Ranges.range(sheet, "B11").getCellFormatText());
	}

	@Test
	public void testBIN2HEX() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("00FB", Ranges.range(sheet, "B13").getCellFormatText());
	}

	@Test
	public void testBIN2OCT() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B15").getCellFormatText());
	}

	@Test
	public void testCOMPLEX() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("3 + 4i", Ranges.range(sheet, "B17").getCellFormatText());
	}

	@Test
	public void testDEC2BIN() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1001", Ranges.range(sheet, "B19").getCellFormatText());
	}

	@Test
	public void testDEC2HEX() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0064", Ranges.range(sheet, "B21").getCellFormatText());
	}

	@Test
	public void testDEC2OCT() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("072", Ranges.range(sheet, "B23").getCellFormatText());
	}

	@Test
	public void testDELTA() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0", Ranges.range(sheet, "B25").getCellFormatText());
	}

	@Test
	public void testERF() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.71", Ranges.range(sheet, "B27").getCellFormatText());
	}

	@Test
	public void testERFC() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.16", Ranges.range(sheet, "B29").getCellFormatText());
	}

	@Test
	public void testGESTEP() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1", Ranges.range(sheet, "B31").getCellFormatText());
	}

	@Test
	public void testHEX2BIN() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("00001111", Ranges.range(sheet, "B33").getCellFormatText());
	}

	@Test
	public void testHEX2DEC() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("165", Ranges.range(sheet, "B35").getCellFormatText());
	}

	@Test
	public void testHEX2OCT() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("017", Ranges.range(sheet, "B37").getCellFormatText());
	}

	@Test
	public void testIMABS() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("13", Ranges.range(sheet, "B39").getCellFormatText());
	}

	@Test
	public void testIMAGINARY() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("4", Ranges.range(sheet, "B41").getCellFormatText());
	}

	@Test
	public void testIMARGUMENT() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.93", Ranges.range(sheet, "B43").getCellFormatText());
	}

	@Test
	public void testIMCONJUNGATE() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("3 - 4i", Ranges.range(sheet, "B45").getCellFormatText());
	}

	@Test
	public void testIMEXP() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1.46869393991589 + 2.28735528717884i", Ranges.range(sheet, "B51").getCellFormatText());
	}

	@Test
	public void testIMDIV() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("5 + 12i", Ranges.range(sheet, "B49").getCellFormatText());
	}

	@Test
	public void testIMPRODUCT() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("27 + 11i", Ranges.range(sheet, "B61").getCellFormatText());
	}

	@Test
	public void testIMREAL() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("6", Ranges.range(sheet, "B63").getCellFormatText());
	}

	@Test
	public void testOCT2BIN() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B73").getCellFormatText());
	}

	@Test
	public void testOCT2DEC() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("44", Ranges.range(sheet, "B75").getCellFormatText());
	}

	@Test
	public void testOCT2HEX() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0040", Ranges.range(sheet, "B77").getCellFormatText());
	}

}
