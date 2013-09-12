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
 * @author kuro, Hawk
 */
public class Formula2007Test {

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

	// FIXME
	// @Test
	// public void testNOW() throws IOException {
	// Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
	// Sheet sheet = book.getSheet("formula-datetime");
	// assertEquals("2013/9/11 16:00", Ranges.range(sheet, "B34").getCellData()
	// .getFormatText());

	// }
	// @Test
	// public void testToday() throws IOException {
	// Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
	// Sheet sheet = book.getSheet("formula-datetime");
	// assertEquals("2013/9/11", Ranges.range(sheet, "B42").getCellData()
	// .getFormatText());
	// }

	// Financial --------------------------------------------------------------------------------
	@Test
	public void testACCRINT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("16.67", Ranges.range(sheet, "B3").getCellData()
				.getFormatText());
	}

	@Test
	public void testACCRINTM() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("20.55", Ranges.range(sheet, "B6").getCellData()
				.getFormatText());
	}

	@Test
	public void testAMORDEGRC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("776", Ranges.range(sheet, "B9").getCellData()
				.getFormatText());
	}

	@Test
	public void testAMORLINC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("360", Ranges.range(sheet, "B12").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOUPDAYBS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("71", Ranges.range(sheet, "B15").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOUPDAYS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("181", Ranges.range(sheet, "B18").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOUPDAYSNC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("110", Ranges.range(sheet, "B21").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOUPNCD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2011/5/15", Ranges.range(sheet, "B24").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOUPNUM() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("4", Ranges.range(sheet, "B27").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOUPPCD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2006/11/15", Ranges.range(sheet, "B30").getCellData()
				.getFormatText());
	}

	@Test
	public void testCUMIPMT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-11135.23", Ranges.range(sheet, "B33").getCellData()
				.getFormatText());
	}

	@Test
	public void testCUMPRINC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-934.11", Ranges.range(sheet, "B36").getCellData()
				.getFormatText());
	}

	@Test
	public void testDB() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("186083.33", Ranges.range(sheet, "B39").getCellData()
				.getFormatText());
	}

	@Test
	public void testDDB() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.32", Ranges.range(sheet, "B42").getCellData()
				.getFormatText());
	}

	@Test
	public void testDISC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B45").getCellData()
				.getFormatText());
	}

	@Test
	public void testDOLLARDE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.125", Ranges.range(sheet, "B47").getCellData()
				.getFormatText());
	}

	@Test
	public void testDOLLARFR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.02", Ranges.range(sheet, "B49").getCellData()
				.getFormatText());
	}

	@Test
	public void testDURATION() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("5.99", Ranges.range(sheet, "B51").getCellData()
				.getFormatText());
	}

	@Test
	public void testEFFECT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B54").getCellData()
				.getFormatText());
	}

	@Test
	public void testFV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2581.40", Ranges.range(sheet, "B57").getCellData()
				.getFormatText());
	}

	@Test
	public void testFVSCHEDULE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.33", Ranges.range(sheet, "B59").getCellData()
				.getFormatText());
	}

	@Test
	public void testINTRATE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B62").getCellData()
				.getFormatText());
	}

	@Test
	public void testIPMT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-292.45", Ranges.range(sheet, "B64").getCellData()
				.getFormatText());
	}

	@Test
	public void testIRR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-2.1%", Ranges.range(sheet, "B67").getCellData()
				.getFormatText());
	}

	@Test
	public void testNOMINAL() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B81").getCellData()
				.getFormatText());
	}

	@Test
	public void testNPER() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("59.67", Ranges.range(sheet, "B84").getCellData()
				.getFormatText());
	}

	@Test
	public void testNPV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("41922.06", Ranges.range(sheet, "B87").getCellData()
				.getFormatText());
	}

	@Test
	public void testPMT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-1037.03", Ranges.range(sheet, "B89").getCellData()
				.getFormatText());
	}

	@Test
	public void testPPMT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-75.62", Ranges.range(sheet, "B91").getCellData()
				.getFormatText());
	}

	@Test
	public void testPRICE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("94.63", Ranges.range(sheet, "B93").getCellData()
				.getFormatText());

	}

	@Test
	public void testPRICEDISC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("99.80", Ranges.range(sheet, "B95").getCellData()
				.getFormatText());
	}

	@Test
	public void testPRICEMAT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("99.98", Ranges.range(sheet, "B97").getCellData()
				.getFormatText());
	}

	@Test
	public void testPV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-59777.15", Ranges.range(sheet, "B99").getCellData()
				.getFormatText());
	}

	@Test
	public void testRATE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1%", Ranges.range(sheet, "B102").getCellData()
				.getFormatText());
	}

	@Test
	public void testRECEIVED() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1014584.65", Ranges.range(sheet, "B104").getCellData()
				.getFormatText());
	}

	@Test
	public void testSLN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2250", Ranges.range(sheet, "B106").getCellData()
				.getFormatText());
	}

	@Test
	public void testSYD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("4090.91", Ranges.range(sheet, "B108").getCellData()
				.getFormatText());
	}

	@Test
	public void testTBILLEQ() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.09", Ranges.range(sheet, "B110").getCellData()
				.getFormatText());
	}

	@Test
	public void testTBILLPRICE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("98.45", Ranges.range(sheet, "B112").getCellData()
				.getFormatText());
	}

	@Test
	public void testTBILLYIELD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.09", Ranges.range(sheet, "B115").getCellData()
				.getFormatText());
	}

	@Test
	public void testXNPV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2086.65", Ranges.range(sheet, "B119").getCellData()
				.getFormatText());
	}

	@Test
	public void testYIELD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B123").getCellData()
				.getFormatText());
	}

	@Test
	public void testYIELDDISC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B125").getCellData()
				.getFormatText());
	}

	@Test
	public void testYIELDMAT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B127").getCellData()
				.getFormatText());
	}

	// Logical --------------------------------------------------------------------------------
	@Test
	public void testAND() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B4").getCellData()
				.getFormatText());
	}

	@Test
	public void testFALSE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("FALSE", Ranges.range(sheet, "B6").getCellData()
				.getFormatText());
	}

	@Test
	public void testIF() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("over 10", Ranges.range(sheet, "B8").getCellData()
				.getFormatText());
	}

	@Test
	public void testIFERROR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("6", Ranges.range(sheet, "B11").getCellData()
				.getFormatText());
	}

	@Test
	public void testNOT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B14").getCellData()
				.getFormatText());
	}

	@Test
	public void testOR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("FALSE", Ranges.range(sheet, "B16").getCellData()
				.getFormatText());
	}

	@Test
	public void testTRUE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B18").getCellData()
				.getFormatText());
	}

	// Math --------------------------------------------------------------------------------
	@Test
	public void testABS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B4").getCellData()
				.getFormatText());
	}

	@Test
	public void testACOS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2.09", Ranges.range(sheet, "B7").getCellData()
				.getFormatText());
	}

	@Test
	public void testACOSH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0", Ranges.range(sheet, "B9").getCellData()
				.getFormatText());
	}

	@Test
	public void testASIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-0.5236", Ranges.range(sheet, "B11").getCellData()
				.getFormatText());
	}

	@Test
	public void testASINH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-1.65", Ranges.range(sheet, "B13").getCellData()
				.getFormatText());
	}

	@Test
	public void testATAN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.79", Ranges.range(sheet, "B15").getCellData()
				.getFormatText());
	}

	@Test
	public void testATAN2() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.79", Ranges.range(sheet, "B17").getCellData()
				.getFormatText());
	}

	@Test
	public void testATANH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1.00", Ranges.range(sheet, "B19").getCellData()
				.getFormatText());
	}

	@Test
	public void testCEILING() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B21").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOMBIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("28", Ranges.range(sheet, "B23").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.50", Ranges.range(sheet, "B25").getCellData()
				.getFormatText());
	}

	@Test
	public void testCOSH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("27.31", Ranges.range(sheet, "B27").getCellData()
				.getFormatText());
	}

	@Test
	public void testDEGREES() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("180", Ranges.range(sheet, "B29").getCellData()
				.getFormatText());
	}

	@Test
	public void testEVEN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B31").getCellData()
				.getFormatText());
	}

	@Test
	public void testEXP() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2.72", Ranges.range(sheet, "B33").getCellData()
				.getFormatText());
	}

	@Test
	public void testFACT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("120", Ranges.range(sheet, "B35").getCellData()
				.getFormatText());
	}

	@Test
	public void testFACTDOUBLE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("48", Ranges.range(sheet, "B37").getCellData()
				.getFormatText());
	}

	@Test
	public void testFLOOR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B39").getCellData()
				.getFormatText());
	}

	@Test
	public void testGCD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B41").getCellData()
				.getFormatText());
	}

	@Test
	public void testINT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("8", Ranges.range(sheet, "B43").getCellData()
				.getFormatText());
	}

	@Test
	public void testLCM() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("10", Ranges.range(sheet, "B45").getCellData()
				.getFormatText());
	}

	@Test
	public void testLN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4.4543", Ranges.range(sheet, "B47").getCellData()
				.getFormatText());
	}

	@Test
	public void testLOG() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B49").getCellData()
				.getFormatText());
	}

	@Test
	public void testLOG10() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B51").getCellData()
				.getFormatText());
	}

	@Test
	public void testMDETERM() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("88", Ranges.range(sheet, "B53").getCellData()
				.getFormatText());
	}

	@Test
	public void testMINVERSE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0", Ranges.range(sheet, "B59").getCellData()
				.getFormatText());
	}

	@Test
	public void testMMULT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B63").getCellData()
				.getFormatText());
	}

	@Test
	public void testMOD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B71").getCellData()
				.getFormatText());
	}

	@Test
	public void testMROUND() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("9", Ranges.range(sheet, "B73").getCellData()
				.getFormatText());
	}

	@Test
	public void testMULTINOMIA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1260", Ranges.range(sheet, "B75").getCellData()
				.getFormatText());
	}

	@Test
	public void testODD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B77").getCellData()
				.getFormatText());
	}

	@Test
	public void testPI() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3.1416", Ranges.range(sheet, "B79").getCellData()
				.getFormatText());
	}

	@Test
	public void testPOWER() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("25", Ranges.range(sheet, "B81").getCellData()
				.getFormatText());
	}

	@Test
	public void testPRODUCT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2250", Ranges.range(sheet, "B83").getCellData()
				.getFormatText());
	}

	@Test
	public void testQUOTIENT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B88").getCellData()
				.getFormatText());
	}

	@Test
	public void testRADIANS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4.7124", Ranges.range(sheet, "B90").getCellData()
				.getFormatText());
	}

	@Test
	public void testROMAN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("CDXCIX", Ranges.range(sheet, "B96").getCellData()
				.getFormatText());
	}

	@Test
	public void testROUND() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2.2", Ranges.range(sheet, "B98").getCellData()
				.getFormatText());
	}

	@Test
	public void testROUNDDOWN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B100").getCellData()
				.getFormatText());
	}

	@Test
	public void testROUNDUP() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B102").getCellData()
				.getFormatText());
	}

	@Test
	public void testSIGN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B111").getCellData()
				.getFormatText());
	}

	@Test
	public void testSIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.00", Ranges.range(sheet, "B113").getCellData()
				.getFormatText());
	}

	@Test
	public void testSINH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1.1752", Ranges.range(sheet, "B115").getCellData()
				.getFormatText());
	}

	@Test
	public void testSQRT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B117").getCellData()
				.getFormatText());
	}

	@Test
	public void testSQRTPI() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1.7725", Ranges.range(sheet, "B119").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUBTOTAL() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("303", Ranges.range(sheet, "B121").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUM() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("5", Ranges.range(sheet, "B127").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUMIF() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("63000", Ranges.range(sheet, "B129").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUMIFS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("20", Ranges.range(sheet, "B135").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUMPRODUCT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("156", Ranges.range(sheet, "B145").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUMSQ() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("25", Ranges.range(sheet, "B150").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUMX2MY2() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-55", Ranges.range(sheet, "B153").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUMX2PY2() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("521", Ranges.range(sheet, "B162").getCellData()
				.getFormatText());
	}

	@Test
	public void testSUMXMY2() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("79", Ranges.range(sheet, "B171").getCellData()
				.getFormatText());
	}

	@Test
	public void testTAN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.9992", Ranges.range(sheet, "B180").getCellData()
				.getFormatText());
	}

	@Test
	public void testTANH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-0.96", Ranges.range(sheet, "B182").getCellData()
				.getFormatText());
	}

	@Test
	public void testTRUNC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("8", Ranges.range(sheet, "B184").getCellData()
				.getFormatText());
	}

	// Date & Time --------------------------------------------------------------------------------
	@Test
	public void testDATE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		CellOperationUtil.applyDataFormat(Ranges.range(sheet, "B3"), "Y/M/D");
		assertEquals("08/1/1", Ranges.range(sheet, "B3").getCellData()
				.getFormatText());// '08/01/01'// is correct and fixed in zpoi
									// 3.9
	}

	@Test
	public void testDATEVALUE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("39448", Ranges.range(sheet, "B6").getCellData()
				.getFormatText());
	}

	@Test
	public void testDAY() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("15", Ranges.range(sheet, "B8").getCellData()
				.getFormatText());
	}

	@Test
	public void testDAY360() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B10").getCellData()
				.getFormatText());
	}

	@Test
	public void testMINUTE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("48", Ranges.range(sheet, "B15").getCellData()
				.getFormatText());
	}

	@Test
	public void testMONTH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("4", Ranges.range(sheet, "B17").getCellData()
				.getFormatText());
	}

	@Test
	public void testNETWORKDAYS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("22", Ranges.range(sheet, "B19").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2013,4,1),DATE(2013,4,30))
		assertEquals("108", Ranges.range(sheet, "B20").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2008,10,1),
									// DATE(2009,3,1))
	}

	@Test
	public void testNetworkDaysStartDateIsHoliday() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("5", Ranges.range(sheet, "B21").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
									// DATE(2013,6,9))
	}

	@Test
	public void testNetworkDaysAllHoliday() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B22").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
									// DATE(2013,6,1))
		assertEquals("0", Ranges.range(sheet, "B23").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
									// DATE(2013,6,2))
	}

	@Test
	public void testNetworkDaysSpecificHoliday() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("107", Ranges.range(sheet, "B24").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2008,10,1),
									// DATE(2009,3,1), DATE(2008,11,26))
		assertEquals("1", Ranges.range(sheet, "B25").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
									// DATE(2013,6,4), DATE(2013,6,3))
	}

	@Test
	public void testStartDateEqualsEndDate() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B26").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2013,6,28),
									// DATE(2013,6,28))
	}

	@Test
	public void testStartDateLaterThanEndDate() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("#VALUE!", Ranges.range(sheet, "B27").getCellData()
				.getFormatText()); // =NETWORKDAYS(DATE(2013,6,2),
									// DATE(2013,6,1))
	}

	@Test
	public void testEmptyDate() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B28").getCellData()
				.getFormatText()); // =NETWORKDAYS(E28,F28)
		// E28 and F28 are blank
	}

	@Test
	public void testSECOND() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("18", Ranges.range(sheet, "B36").getCellData()
				.getFormatText());
	}

	@Test
	public void testTIME() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0.5", Ranges.range(sheet, "B38").getCellData()
				.getFormatText());
	}

	@Test
	public void testWEEKDAY() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("5", Ranges.range(sheet, "B44").getCellData()
				.getFormatText());
	}

	@Test
	public void testWORKDAY() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/4/5", Ranges.range(sheet, "B46").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),4)
		assertEquals("2009/4/30", Ranges.range(sheet, "B47").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2008,10,1),151)
	}

	@Test
	public void testStartDateIsHoliday() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/6/3", Ranges.range(sheet, "B48").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,1),1)
		assertEquals("2013/5/31", Ranges.range(sheet, "B49").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,1), -1)
	}

	@Test
	public void testEndDateIsHoliday() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/4/8", Ranges.range(sheet, "B50").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),5)
		assertEquals("2013/4/8", Ranges.range(sheet, "B51").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,5),1)
	}

	@Test
	public void testNegativeWorkdayEndDateIsHolday() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/3/29", Ranges.range(sheet, "B52").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),-1)
		assertEquals("2013/5/31", Ranges.range(sheet, "B53").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,7),-5)
	}

	@Test
	public void testWorkdayBoundary() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/4/1", Ranges.range(sheet, "B54").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),0)
		assertEquals("2013/6/1", Ranges.range(sheet, "B55").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,1),0)
	}

	@Test
	public void testWorkdaySpecifiedholiday() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/4/5", Ranges.range(sheet, "B56").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),3, DATE(2013,4,2))
		assertEquals("2013/4/3", Ranges.range(sheet, "B57").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,7),
									// -3,DATE(2013,4,2))
		assertEquals("2013/4/1", Ranges.range(sheet, "B58").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,3),
									// -1,DATE(2013,4,2))
	}

	@Test
	public void testYEAR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2008", Ranges.range(sheet, "B61").getCellData()
				.getFormatText()); // =YEAR(DATE(2008,1,1))
	}

	@Test
	public void testYEARFRAC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals(0.58, Ranges.range(sheet, "B63").getCellData()
				.getDoubleValue(), 0.005); // =YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))
	}
	
	// Text ---------------------------------------------------------------------------------------
	@Test
	public void testCHAR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("A", Ranges.range(sheet, "B4").getCellFormatText());
	}

	@Test
	public void testLOWER() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("e. e. cummings", Ranges.range(sheet, "B34")
				.getCellFormatText());
	}

	@Test
	public void testCODE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("65", Ranges.range(sheet, "B9").getCellFormatText());
	}

	@Test
	public void testCLEAN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("text", Ranges.range(sheet, "B6").getCellFormatText());
	}

	@Test
	public void testCONCATENATE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("ZK", Ranges.range(sheet, "B11").getCellFormatText());
	}

	@Test
	public void testEXACT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("TRUE", Ranges.range(sheet, "B15").getCellFormatText());
	}

	@Test
	public void testFIND() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B19").getCellFormatText());
	}

	@Test
	public void testFIXED() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1,234.6", Ranges.range(sheet, "B23").getCellFormatText());
	}

	@Test
	public void testLEFT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("\u6771 \u4EAC ", Ranges.range(sheet, "B26")
				.getCellFormatText());
	}

	@Test
	public void testLEN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("11", Ranges.range(sheet, "B30").getCellFormatText());
	}

	@Test
	public void testMID() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Fluid", Ranges.range(sheet, "B37").getCellFormatText());
	}

	@Test
	public void testPROPER() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("This Is A Title", Ranges.range(sheet, "B41")
				.getCellFormatText());
	}

	@Test
	public void testREPLACE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("abcde*k", Ranges.range(sheet, "B44").getCellFormatText());
	}

	@Test
	public void testREPT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("*-*-*-", Ranges.range(sheet, "B48").getCellFormatText());
	}

	@Test
	public void testRIGHT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Price", Ranges.range(sheet, "B50").getCellFormatText());
	}

	@Test
	public void testSEARCH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("6", Ranges.range(sheet, "B54").getCellFormatText());
	}

	@Test
	public void testSUBSTITUTE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Quarter 2, 2008", Ranges.range(sheet, "B58")
				.getCellFormatText());
	}

	@Test
	public void testT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Sale Price", Ranges.range(sheet, "B62")
				.getCellFormatText());
	}

	@Test
	public void testTEXT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Date: 2007-08-06", Ranges.range(sheet, "B64")
				.getCellFormatText());
	}

	@Test
	public void testTRIM() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("revenue in quarter 1", Ranges.range(sheet, "B67")
				.getCellFormatText());
	}

	@Test
	public void testUPPER() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("TOTAL", Ranges.range(sheet, "B69").getCellFormatText());
	}

	//Info ----------------------------------------------------------------------------------------
	@Test
	public void testERRORTYPE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("1", Ranges.range(sheet, "B3").getCellFormatText());
	}

	@Test
	public void testISBLANK() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B10").getCellFormatText());
	}

	@Test
	public void testISERR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B12").getCellFormatText());
	}

	@Test
	public void testISERROR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B15").getCellFormatText());
	}

	@Test
	public void testISEVEN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B18").getCellFormatText());
	}

	@Test
	public void testISLOGICAL() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B21").getCellFormatText());
	}

	@Test
	public void testISNA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B23").getCellFormatText());
	}

	@Test
	public void testISNONTEXT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B26").getCellFormatText());
	}

	@Test
	public void testISNUMBER() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B29").getCellFormatText());
	}

	@Test
	public void testISODD() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B31").getCellFormatText());
	}

	@Test
	public void testISREF() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B33").getCellFormatText());
	}

	@Test
	public void testISTEXT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B35").getCellFormatText());
	}

	@Test
	public void testN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("7", Ranges.range(sheet, "B38").getCellFormatText());
	}

	@Test
	public void testNA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("#N/A", Ranges.range(sheet, "B41").getCellFormatText());
	}

	@Test
	public void testTYPE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("2", Ranges.range(sheet, "B43").getCellFormatText());
	}

	// statistical ----------------------------------------------------------------------------------------
	@Test
	public void testAVEDEV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.02", Ranges.range(sheet, "B3").getCellFormatText());
	}

	@Test
	public void testAVERAGE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("11", Ranges.range(sheet, "B6").getCellFormatText());
	}	

	@Test
	public void testAVERAGEA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.6", Ranges.range(sheet, "B8").getCellFormatText());
	}

	@Test
	public void testBIOMDIST() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.21", Ranges.range(sheet, "B21").getCellFormatText());
	}

	@Test
	public void testCHIDIST() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.05", Ranges.range(sheet, "B23").getCellFormatText());
	}

	@Test
	public void testCHIINV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("18.31", Ranges.range(sheet, "B25").getCellFormatText());
	}

	@Test
	public void testCOUNT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B41").getCellFormatText());
	}

	@Test
	public void testCOUNTA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("6", Ranges.range(sheet, "B43").getCellFormatText());
	}

	@Test
	public void testCOUNTBLANK() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B45").getCellFormatText());
	}

	@Test
	public void testCOUNTIF() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B47").getCellFormatText());
	}

	@Test
	public void testDEVSQ() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("48", Ranges.range(sheet, "B55").getCellFormatText());
	}

	@Test
	public void testEXPONDIST() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.86", Ranges.range(sheet, "B57").getCellFormatText());
	}

	@Test
	public void testGAMMAINV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("10.00", Ranges.range(sheet, "B78").getCellFormatText());
	}

	@Test
	public void testGAMMALN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.79", Ranges.range(sheet, "B80").getCellFormatText());
	}

	@Test
	public void testGEOMEAN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.48", Ranges.range(sheet, "B82").getCellFormatText());
	}

	@Test
	public void testHYPGEOMDIST() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.36", Ranges.range(sheet, "B92").getCellFormatText());
	}

	@Test
	public void testKURT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("-0.15", Ranges.range(sheet, "B97").getCellFormatText());
	}

	@Test
	public void testLARGE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5", Ranges.range(sheet, "B100").getCellFormatText());
	}

	@Test
	public void testMAX() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("27", Ranges.range(sheet, "B110").getCellFormatText());
	}

	@Test
	public void testMAXA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.5", Ranges.range(sheet, "B112").getCellFormatText());
	}

	@Test
	public void testMEDIAN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B114").getCellFormatText());
	}

	@Test
	public void testMIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B116").getCellFormatText());
	}

	@Test
	public void testMINA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0", Ranges.range(sheet, "B118").getCellFormatText());
	}

	@Test
	public void testMODE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B121").getCellFormatText());
	}

	@Test
	public void testNORMDIST() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.91", Ranges.range(sheet, "B125").getCellFormatText());
	}

	@Test
	public void testPOISSON() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.12", Ranges.range(sheet, "B142").getCellFormatText());
	}

	@Test
	public void testRANK() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B149").getCellFormatText());
	}

	@Test
	public void testSKEW() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.36", Ranges.range(sheet, "B154").getCellFormatText());
	}

	@Test
	public void testSLOPE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.31", Ranges.range(sheet, "B156").getCellFormatText());
	}

	@Test
	public void testSMALL() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B159").getCellFormatText());
	}

	@Test
	public void testSTDEV() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("27.46", Ranges.range(sheet, "B164").getCellFormatText());
	}

	@Test
	public void testVAR() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("754.27", Ranges.range(sheet, "B189").getCellFormatText());
	}

	@Test
	public void testVARP() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("678.84", Ranges.range(sheet, "B193").getCellFormatText());
	}

	@Test
	public void testWEIBULL() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.93", Ranges.range(sheet, "B197").getCellFormatText());
	}

	// engineering  ----------------------------------------------------------------------------------------
	@Test
	public void testBESSELI() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.98", Ranges.range(sheet, "B3").getCellFormatText());
	}

	@Test
	public void testBESSELJ() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.33", Ranges.range(sheet, "B5").getCellFormatText());
	}

	@Test
	public void testBESSELK() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.28", Ranges.range(sheet, "B7").getCellFormatText());
	}

	@Test
	public void testBESSELY() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.15", Ranges.range(sheet, "B9").getCellFormatText());
	}

	@Test
	public void testBIN2DEC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("100", Ranges.range(sheet, "B11").getCellFormatText());
	}

	@Test
	public void testBIN2HEX() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("00FB", Ranges.range(sheet, "B13").getCellFormatText());
	}

	@Test
	public void testBIN2OCT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B15").getCellFormatText());
	}

	@Test
	public void testDEC2BIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1001", Ranges.range(sheet, "B19").getCellFormatText());
	}

	@Test
	public void testDEC2HEX() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0064", Ranges.range(sheet, "B21").getCellFormatText());
	}

	@Test
	public void testDEC2OCT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("072", Ranges.range(sheet, "B23").getCellFormatText());
	}

	@Test
	public void testDELTA() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0", Ranges.range(sheet, "B25").getCellFormatText());
	}

	@Test
	public void testERF() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.71", Ranges.range(sheet, "B27").getCellFormatText());
	}

	@Test
	public void testERFC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.16", Ranges.range(sheet, "B29").getCellFormatText());
	}

	@Test
	public void testGESTEP() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1", Ranges.range(sheet, "B31").getCellFormatText());
	}

	@Test
	public void testHEX2BIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("00001111", Ranges.range(sheet, "B33").getCellFormatText());
	}

	@Test
	public void testHEX2DEC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("165", Ranges.range(sheet, "B35").getCellFormatText());
	}

	@Test
	public void testHEX2OCT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("017", Ranges.range(sheet, "B37").getCellFormatText());
	}

	@Test
	public void testIMABS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("13", Ranges.range(sheet, "B39").getCellFormatText());
	}

	@Test
	public void testIMAGINARY() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("4", Ranges.range(sheet, "B41").getCellFormatText());
	}

	@Test
	public void testIMARGUMENT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.93", Ranges.range(sheet, "B43").getCellFormatText());
	}

	@Test
	public void testIMREAL() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("6", Ranges.range(sheet, "B63").getCellFormatText());
	}

	@Test
	public void testOCT2BIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B73").getCellFormatText());
	}

	@Test
	public void testOCT2DEC() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("44", Ranges.range(sheet, "B75").getCellFormatText());
	}

	@Test
	public void testOCT2HEX() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0040", Ranges.range(sheet, "B77").getCellFormatText());
	}
	
	@Test
	public void testIMCOS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("0.833730025131149-0.988897705762865i", Ranges.range(sheet, "B47").getCellFormatText());
	}

	@Test
	public void testIMLN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("1.6094379124341+0.927295218001612i", Ranges.range(sheet, "B53").getCellFormatText());
		
	}

	@Test
	public void testIMSQRT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("1.09868411346781+0.455089860562227i", Ranges.range(sheet, "B67").getCellFormatText());
		
	}

	@Test
	public void testIMLOG10() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("0.698970004336019+0.402719196273373i", Ranges.range(sheet, "B55").getCellFormatText());
		
	}

	@Test
	public void testIMSUM() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("8+i", Ranges.range(sheet, "B71").getCellFormatText());
		
	}

	@Test
	public void testIMSIN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("3.85373803791938-27.0168132580039i", Ranges.range(sheet, "B65").getCellFormatText());
		
	}

	@Test
	public void testIMSUB() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("8+i", Ranges.range(sheet, "B69").getCellFormatText());
		
	}

	@Test
	public void testIMLOG2() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("2.32192809506607+1.33780421255394i", Ranges.range(sheet, "B57").getCellFormatText());
		
	}

	@Test
	public void testIMPOWER() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("-46+9.00000000000001i", Ranges.range(sheet, "B59").getCellFormatText());
	}

	// lookup  ----------------------------------------------------------------------------------------
	@Test
	public void testADDRESS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("$C$2", Ranges.range(sheet, "B3").getCellFormatText());
	}
	
	@Test
	public void testCHOOSE() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2nd", Ranges.range(sheet, "B7").getCellFormatText());
	}
	
	@Test
	public void testCOLUMN() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2", Ranges.range(sheet, "B10").getCellFormatText());
	}
	
	@Test
	public void testCOLUMNS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("3", Ranges.range(sheet, "B12").getCellFormatText());
	}
	
	@Test
	public void testHLOOKUP() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("4", Ranges.range(sheet, "B14").getCellFormatText());
	}
	
	@Test
	public void testHyperLink() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("ZK", Ranges.range(sheet, "B20").getCellFormatText());
	}

	@Test
	public void testINDIRECT() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("1.333", Ranges.range(sheet, "B26").getCellFormatText());
	}

	@Test
	public void testLOOKUP() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("orange", Ranges.range(sheet, "B29").getCellFormatText());
	}

	@Test
	public void testMATCH() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2", Ranges.range(sheet, "B36").getCellFormatText());
	}

	@Test
	public void testOFFSET() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("offset input", Ranges.range(sheet, "B42").getCellFormatText());
	}

	@Test
	public void testROW() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("45", Ranges.range(sheet, "B45").getCellFormatText());
	}

	@Test
	public void testROWS() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("4", Ranges.range(sheet, "B48").getCellFormatText());
	}

	@Test
	public void testVLOOKUP() throws IOException {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2.93", Ranges.range(sheet, "B56").getCellFormatText());
	}

}
