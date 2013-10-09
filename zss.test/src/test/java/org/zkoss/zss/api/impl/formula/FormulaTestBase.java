package org.zkoss.zss.api.impl.formula;

import static org.junit.Assert.assertEquals;

import org.zkoss.zss.AssertUtil;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * all formula assert utility.
 * 
 * @author kuro
 */
public class FormulaTestBase {
	
	// FIXME
	// 
	// protected void testNOW()  {
	// Book book = Util.loadBook("book/TestFile2003-Formula.xls");
	// Sheet sheet = book.getSheet("formula-datetime");
	// assertEquals("2013/9/11 16:00", Ranges.range(sheet, "B34").getCellData()
	// .getFormatText());
	// }
	// 
	// protected void testToday()  {
	// Book book = Util.loadBook("book/TestFile2003-Formula.xls");
	// Sheet sheet = book.getSheet("formula-datetime");
	// assertEquals("2013/9/11", Ranges.range(sheet, "B42").getCellData()
	// .getFormatText());
	// }
	
	// spreadsheet doesn't change format into currency automatically.
	protected void testDOLLAR(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("$1,234.57", Ranges.range(sheet, "B13").getCellFormatText());
	}

	protected void testACCRINT(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("16.67", Ranges.range(sheet, "B3").getCellData().getFormatText());
	}

	
	protected void testACCRINTM(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("20.55", Ranges.range(sheet, "B6").getCellData().getFormatText());
	}

	
	protected void testAMORDEGRC(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("776", Ranges.range(sheet, "B9").getCellData().getFormatText());
	}

	
	protected void testAMORLINC(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("360", Ranges.range(sheet, "B12").getCellData().getFormatText());
	}

	
	protected void testCOUPDAYBS(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("71", Ranges.range(sheet, "B15").getCellData().getFormatText());
	}

	
	protected void testCOUPDAYS(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("181", Ranges.range(sheet, "B18").getCellData().getFormatText());
	}

	
	protected void testCOUPDAYSNC(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("110", Ranges.range(sheet, "B21").getCellData().getFormatText());
	}
	
	protected void testCOUPNCD(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2011/5/15", Ranges.range(sheet, "B24").getCellData().getFormatText());
	}

	protected void testCOUPPCD(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2006/11/15", Ranges.range(sheet, "B30").getCellData().getFormatText());
	}

	
	protected void testCOUPNUM(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("4", Ranges.range(sheet, "B27").getCellData().getFormatText());
	}

	
	protected void testCUMIPMT(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-11135.23", Ranges.range(sheet, "B33").getCellData().getFormatText());
	}

	
	protected void testCUMPRINC(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-934.11", Ranges.range(sheet, "B36").getCellData().getFormatText());
	}

	
	protected void testDB(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("186083.33", Ranges.range(sheet, "B39").getCellData().getFormatText());
	}

	
	protected void testDDB(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.32", Ranges.range(sheet, "B42").getCellData().getFormatText());
	}

	protected void testDISC(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B45").getCellData().getFormatText());
	}

	
	protected void testDOLLARDE(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.125", Ranges.range(sheet, "B47").getCellData().getFormatText());
	}

	
	protected void testDOLLARFR(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.02", Ranges.range(sheet, "B49").getCellData().getFormatText());
	}

	
	protected void testDURATION(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("5.99", Ranges.range(sheet, "B51").getCellData().getFormatText());
	}

	
	protected void testEFFECT(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B54").getCellData().getFormatText());
	}

	
	protected void testFV(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2581.40", Ranges.range(sheet, "B57").getCellData().getFormatText());
	}

	
	protected void testFVSCHEDULE(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.33", Ranges.range(sheet, "B59").getCellData().getFormatText());
	}

	
	protected void testINTRATE(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B62").getCellData().getFormatText());
	}

	
	protected void testIPMT(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-292.45", Ranges.range(sheet, "B64").getCellData().getFormatText());
	}

	
	protected void testIRR(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-2.1%", Ranges.range(sheet, "B67").getCellData().getFormatText());
	}

	
	protected void testNOMINAL(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B81").getCellData().getFormatText());
	}

	
	protected void testNPER(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("59.67", Ranges.range(sheet, "B84").getCellData().getFormatText());
	}

	
	protected void testNPV(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("41922.06", Ranges.range(sheet, "B87").getCellData().getFormatText());
	}

	
	protected void testPMT(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-1037.03", Ranges.range(sheet, "B89").getCellData().getFormatText());
	}

	
	protected void testPPMT(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-75.62", Ranges.range(sheet, "B91").getCellData().getFormatText());
	}

	
	protected void testPRICE(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("94.63", Ranges.range(sheet, "B93").getCellData().getFormatText());

	}

	
	protected void testPRICEDISC(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("99.80", Ranges.range(sheet, "B95").getCellData().getFormatText());
	}

	
	protected void testPRICEMAT(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("99.98", Ranges.range(sheet, "B97").getCellData().getFormatText());
	}

	
	protected void testPV(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-59777.15", Ranges.range(sheet, "B99").getCellData().getFormatText());
	}

	
	protected void testRATE(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1%", Ranges.range(sheet, "B102").getCellData().getFormatText());
	}

	
	protected void testRECEIVED(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1014584.65", Ranges.range(sheet, "B104").getCellData().getFormatText());
	}

	
	protected void testSLN(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2250", Ranges.range(sheet, "B106").getCellData().getFormatText());
	}

	
	protected void testSYD(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("4090.91", Ranges.range(sheet, "B108").getCellData().getFormatText());
	}

	
	protected void testTBILLEQ(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.09", Ranges.range(sheet, "B110").getCellData().getFormatText());
	}

	
	protected void testTBILLPRICE(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("98.45", Ranges.range(sheet, "B112").getCellData().getFormatText());
	}

	
	protected void testTBILLYIELD(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.09", Ranges.range(sheet, "B115").getCellData().getFormatText());
	}

	
	protected void testXNPV(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("2086.65", Ranges.range(sheet, "B119").getCellData().getFormatText());
	}

	
	protected void testYIELD(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B123").getCellData().getFormatText());
	}

	
	protected void testYIELDDISC(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B125").getCellData().getFormatText());
	}

	
	protected void testYIELDMAT(Book book) {

		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B127").getCellData().getFormatText());
	}

	
	protected void testAND(Book book) {

		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B4").getCellData().getFormatText());
	}

	
	protected void testFALSE(Book book) {

		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("FALSE", Ranges.range(sheet, "B6").getCellData().getFormatText());
	}

	
	protected void testIF(Book book) {

		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("over 10", Ranges.range(sheet, "B8").getCellData().getFormatText());
	}

	
	protected void testIFERROR(Book book) {

		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("6", Ranges.range(sheet, "B11").getCellData().getFormatText());
	}

	
	protected void testNOT(Book book) {

		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B14").getCellData().getFormatText());
	}

	
	protected void testOR(Book book) {

		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("FALSE", Ranges.range(sheet, "B16").getCellData().getFormatText());
	}

	
	protected void testTRUE(Book book) {

		Sheet sheet = book.getSheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B18").getCellData().getFormatText());
	}

	
	protected void testABS(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B4").getCellData().getFormatText());
	}

	
	protected void testACOS(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2.09", Ranges.range(sheet, "B7").getCellData().getFormatText());
	}

	
	protected void testACOSH(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0", Ranges.range(sheet, "B9").getCellData().getFormatText());
	}

	
	protected void testASIN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-0.5236", Ranges.range(sheet, "B11").getCellData().getFormatText());
	}

	
	protected void testASINH(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-1.65", Ranges.range(sheet, "B13").getCellData().getFormatText());
	}

	
	protected void testATAN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.79", Ranges.range(sheet, "B15").getCellData().getFormatText());
	}

	
	protected void testATAN2(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.79", Ranges.range(sheet, "B17").getCellData().getFormatText());
	}

	
	protected void testATANH(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1.00", Ranges.range(sheet, "B19").getCellData().getFormatText());
	}

	
	protected void testCEILING(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B21").getCellData().getFormatText());
	}

	
	protected void testCOMBIN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("28", Ranges.range(sheet, "B23").getCellData().getFormatText());
	}

	
	protected void testCOS(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.50", Ranges.range(sheet, "B25").getCellData().getFormatText());
	}

	
	protected void testCOSH(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("27.31", Ranges.range(sheet, "B27").getCellData().getFormatText());
	}

	
	protected void testDEGREES(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("180", Ranges.range(sheet, "B29").getCellData().getFormatText());
	}

	
	protected void testEVEN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B31").getCellData().getFormatText());
	}

	
	protected void testEXP(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2.72", Ranges.range(sheet, "B33").getCellData().getFormatText());
	}

	
	protected void testFACT(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("120", Ranges.range(sheet, "B35").getCellData().getFormatText());
	}

	
	protected void testFACTDOUBLE(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("48", Ranges.range(sheet, "B37").getCellData().getFormatText());
	}

	
	protected void testFLOOR(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B39").getCellData().getFormatText());
	}

	
	protected void testGCD(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B41").getCellData().getFormatText());
	}

	
	protected void testINT(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("8", Ranges.range(sheet, "B43").getCellData().getFormatText());
	}

	
	protected void testLCM(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("10", Ranges.range(sheet, "B45").getCellData().getFormatText());
	}

	
	protected void testLN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4.4543", Ranges.range(sheet, "B47").getCellData().getFormatText());
	}

	
	protected void testLOG(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B49").getCellData().getFormatText());
	}

	
	protected void testLOG10(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B51").getCellData().getFormatText());
	}

	
	protected void testMDETERM(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("88", Ranges.range(sheet, "B53").getCellData().getFormatText());
	}

	
	protected void testMINVERSE(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0", Ranges.range(sheet, "B59").getCellData().getFormatText());
	}

	
	protected void testMMULT(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B63").getCellData().getFormatText());
	}

	
	protected void testMOD(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B71").getCellData().getFormatText());
	}

	
	protected void testMROUND(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("9", Ranges.range(sheet, "B73").getCellData().getFormatText());
	}

	
	protected void testMULTINOMIA(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1260", Ranges.range(sheet, "B75").getCellData().getFormatText());
	}

	
	protected void testODD(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B77").getCellData().getFormatText());
	}

	
	protected void testPI(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3.1416", Ranges.range(sheet, "B79").getCellData().getFormatText());
	}

	
	protected void testPOWER(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("25", Ranges.range(sheet, "B81").getCellData().getFormatText());
	}

	
	protected void testPRODUCT(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2250", Ranges.range(sheet, "B83").getCellData().getFormatText());
	}

	
	protected void testQUOTIENT(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B88").getCellData().getFormatText());
	}

	
	protected void testRADIANS(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4.7124", Ranges.range(sheet, "B90").getCellData().getFormatText());
	}

	
	protected void testROMAN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("CDXCIX", Ranges.range(sheet, "B96").getCellData().getFormatText());
	}

	
	protected void testROUND(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("2.2", Ranges.range(sheet, "B98").getCellData().getFormatText());
	}

	
	protected void testROUNDDOWN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B100").getCellData().getFormatText());
	}

	
	protected void testROUNDUP(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B102").getCellData().getFormatText());
	}

	
	protected void testSIGN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B111").getCellData().getFormatText());
	}

	
	protected void testSIN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.00", Ranges.range(sheet, "B113").getCellData().getFormatText());
	}

	
	protected void testSINH(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1.1752", Ranges.range(sheet, "B115").getCellData().getFormatText());
	}

	
	protected void testSQRT(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B117").getCellData().getFormatText());
	}

	
	protected void testSQRTPI(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("1.7725", Ranges.range(sheet, "B119").getCellData().getFormatText());
	}

	
	protected void testSUBTOTAL(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("303", Ranges.range(sheet, "B121").getCellData().getFormatText());
	}

	
	protected void testSUM(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("5", Ranges.range(sheet, "B127").getCellData().getFormatText());
	}

	
	protected void testSUMIF(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("63000", Ranges.range(sheet, "B129").getCellData().getFormatText());
	}

	
	protected void testSUMIFS(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("20", Ranges.range(sheet, "B135").getCellData().getFormatText());
	}

	
	protected void testSUMPRODUCT(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("156", Ranges.range(sheet, "B145").getCellData().getFormatText());
	}

	
	protected void testSUMSQ(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("25", Ranges.range(sheet, "B150").getCellData().getFormatText());
	}

	
	protected void testSUMX2MY2(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-55", Ranges.range(sheet, "B153").getCellData().getFormatText());
	}

	
	protected void testSUMX2PY2(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("521", Ranges.range(sheet, "B162").getCellData().getFormatText());
	}

	
	protected void testSUMXMY2(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("79", Ranges.range(sheet, "B171").getCellData().getFormatText());
	}

	
	protected void testTAN(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.9992", Ranges.range(sheet, "B180").getCellData().getFormatText());
	}

	
	protected void testTANH(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-0.96", Ranges.range(sheet, "B182").getCellData().getFormatText());
	}

	
	protected void testTRUNC(Book book) {

		Sheet sheet = book.getSheet("formula-math");
		assertEquals("8", Ranges.range(sheet, "B184").getCellData().getFormatText());
	}

	
	protected void testDATE(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		CellOperationUtil.applyDataFormat(Ranges.range(sheet, "B3"), "Y/M/D");
		assertEquals("08/1/1", Ranges.range(sheet, "B3").getCellData().getFormatText());
		// '08/01/01' is correct and fixed in zpoi 3.9.
	}

	
	protected void testDATEVALUE(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("39448", Ranges.range(sheet, "B6").getCellData().getFormatText());
	}

	
	protected void testDAY(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("15", Ranges.range(sheet, "B8").getCellData().getFormatText());
	}

	
	protected void testDAY360(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B10").getCellData().getFormatText());
	}

	
	protected void testHOUR(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("3", Ranges.range(sheet, "B13").getCellData().getFormatText());
	}
	
	protected void testMINUTE(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("48", Ranges.range(sheet, "B15").getCellData().getFormatText());
	}

	
	protected void testMONTH(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("4", Ranges.range(sheet, "B17").getCellData().getFormatText());
	}

	
	protected void testNETWORKDAYS(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("22", Ranges.range(sheet, "B19").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2013,4,1),DATE(2013,4,30))
		assertEquals("108", Ranges.range(sheet, "B20").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2008,10,1),
																						// DATE(2009,3,1))
	}

	
	protected void testNetworkDaysStartDateIsHoliday(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("5", Ranges.range(sheet, "B21").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,9))
	}

	
	protected void testNetworkDaysAllHoliday(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B22").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,1))
		assertEquals("0", Ranges.range(sheet, "B23").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,2))
	}

	
	protected void testNetworkDaysSpecificHoliday(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("107", Ranges.range(sheet, "B24").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2008,10,1),
																						// DATE(2009,3,1),
																						// DATE(2008,11,26))
		assertEquals("1", Ranges.range(sheet, "B25").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,4),
																						// DATE(2013,6,3))
	}

	
	protected void testStartDateEqualsEndDate(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B26").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2013,6,28),
																						// DATE(2013,6,28))
	}

	
	protected void testStartDateLaterThanEndDate(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("#VALUE!", Ranges.range(sheet, "B27").getCellData().getFormatText()); // =NETWORKDAYS(DATE(2013,6,2),
																							// DATE(2013,6,1))
	}

	
	protected void testEmptyDate(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B28").getCellData().getFormatText()); // =NETWORKDAYS(E28,F28)
		// E28 and F28 are blank
	}

	
	protected void testSECOND(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("18", Ranges.range(sheet, "B36").getCellData().getFormatText());
	}

	
	protected void testTIME(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0.5", Ranges.range(sheet, "B38").getCellData().getFormatText());
	}

	
	protected void testWEEKDAY(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("5", Ranges.range(sheet, "B44").getCellData().getFormatText());
	}
	
	protected void testWORKDAY(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/4/5", Ranges.range(sheet, "B46").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),4)
		assertEquals("2009/4/30", Ranges.range(sheet, "B47").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2008,10,1),151)
	}

	protected void testStartDateIsHoliday(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/6/3", Ranges.range(sheet, "B48").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,1),1)
		assertEquals("2013/5/31", Ranges.range(sheet, "B49").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,1), -1)
	}

	protected void testEndDateIsHoliday(Book book)  {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/4/8", Ranges.range(sheet, "B50").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),5)
		assertEquals("2013/4/8", Ranges.range(sheet, "B51").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,5),1)
	}

	protected void testNegativeWorkdayEndDateIsHolday(Book book)  {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/3/29", Ranges.range(sheet, "B52").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),-1)
		assertEquals("2013/5/31", Ranges.range(sheet, "B53").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,7),-5)
	}

	protected void testWorkdayBoundary(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2013/4/1", Ranges.range(sheet, "B54").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,4,1),0)
		assertEquals("2013/6/1", Ranges.range(sheet, "B55").getCellData()
				.getFormatText()); // =WORKDAY(DATE(2013,6,1),0)
	}

	protected void testWorkdaySpecifiedholiday(Book book) {
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

	
	protected void testYEAR(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("2008", Ranges.range(sheet, "B61").getCellData().getFormatText()); // =YEAR(DATE(2008,1,1))
	}

	
	protected void testYEARFRAC(Book book) {

		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals(0.58, Ranges.range(sheet, "B63").getCellData().getDoubleValue(), 0.005); // =YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))
	}

	
	protected void testCHAR(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("A", Ranges.range(sheet, "B4").getCellFormatText());
	}

	
	protected void testLOWER(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("e. e. cummings", Ranges.range(sheet, "B34").getCellFormatText());
	}

	
	protected void testCODE(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("65", Ranges.range(sheet, "B9").getCellFormatText());
	}

	
	protected void testCLEAN(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("text", Ranges.range(sheet, "B6").getCellFormatText());
	}

	
	protected void testCONCATENATE(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("ZK", Ranges.range(sheet, "B11").getCellFormatText());
	}

	
	protected void testEXACT(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("TRUE", Ranges.range(sheet, "B15").getCellFormatText());
	}

	
	protected void testFIND(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B19").getCellFormatText());
	}

	
	protected void testFIXED(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1,234.6", Ranges.range(sheet, "B23").getCellFormatText());
	}

	
	protected void testLEFT(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("\u6771 \u4EAC ", Ranges.range(sheet, "B26").getCellFormatText());
	}

	
	protected void testLEN(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("11", Ranges.range(sheet, "B30").getCellFormatText());
	}

	
	protected void testMID(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Fluid", Ranges.range(sheet, "B37").getCellFormatText());
	}

	
	protected void testPROPER(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("This Is A Title", Ranges.range(sheet, "B41").getCellFormatText());
	}

	
	protected void testREPLACE(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("abcde*k", Ranges.range(sheet, "B44").getCellFormatText());
	}

	
	protected void testREPT(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("*-*-*-", Ranges.range(sheet, "B48").getCellFormatText());
	}

	
	protected void testRIGHT(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Price", Ranges.range(sheet, "B50").getCellFormatText());
	}

	
	protected void testSEARCH(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("6", Ranges.range(sheet, "B54").getCellFormatText());
	}

	
	protected void testSUBSTITUTE(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Quarter 2, 2008", Ranges.range(sheet, "B58").getCellFormatText());
	}

	
	protected void testT(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Sale Price", Ranges.range(sheet, "B62").getCellFormatText());
	}

	
	protected void testTEXT(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Date: 2007-08-06", Ranges.range(sheet, "B64").getCellFormatText());
	}

	
	protected void testTRIM(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("revenue in quarter 1", Ranges.range(sheet, "B67").getCellFormatText());
	}

	
	protected void testUPPER(Book book) {

		Sheet sheet = book.getSheet("formula-text");
		assertEquals("TOTAL", Ranges.range(sheet, "B69").getCellFormatText());
	}

	
	protected void testAVEDEV(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.02", Ranges.range(sheet, "B3").getCellFormatText());
	}

	
	protected void testAVERAGE(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("11", Ranges.range(sheet, "B6").getCellFormatText());
	}

	
	protected void testERRORTYPE(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("1", Ranges.range(sheet, "B3").getCellFormatText());
	}

	
	protected void testISBLANK(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B10").getCellFormatText());
	}

	
	protected void testISERR(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B12").getCellFormatText());
	}

	
	protected void testISERROR(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B15").getCellFormatText());
	}

	
	protected void testISEVEN(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B18").getCellFormatText());
	}

	
	protected void testISLOGICAL(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B21").getCellFormatText());
	}

	
	protected void testISNA(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B23").getCellFormatText());
	}

	
	protected void testISNONTEXT(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B26").getCellFormatText());
	}

	
	protected void testISNUMBER(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B29").getCellFormatText());
	}

	
	protected void testISODD(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B31").getCellFormatText());
	}

	
	protected void testISREF(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B33").getCellFormatText());
	}

	
	protected void testISTEXT(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B35").getCellFormatText());
	}

	
	protected void testN(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("7", Ranges.range(sheet, "B38").getCellFormatText());
	}

	
	protected void testNA(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("#N/A", Ranges.range(sheet, "B41").getCellFormatText());
	}

	
	protected void testTYPE(Book book) {

		Sheet sheet = book.getSheet("formula-info");
		assertEquals("2", Ranges.range(sheet, "B43").getCellFormatText());
	}

	
	protected void testAVERAGEA(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.6", Ranges.range(sheet, "B8").getCellFormatText());
	}

	
	protected void testBINOMDIST(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.21", Ranges.range(sheet, "B21").getCellFormatText());
	}

	
	protected void testCHIDIST(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.05", Ranges.range(sheet, "B23").getCellFormatText());
	}

	
	protected void testCHIINV(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("18.31", Ranges.range(sheet, "B25").getCellFormatText());
	}

	
	protected void testCOUNT(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B41").getCellFormatText());
	}

	
	protected void testCOUNTA(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("6", Ranges.range(sheet, "B43").getCellFormatText());
	}

	
	protected void testCOUNTBLANK(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B45").getCellFormatText());
	}

	
	protected void testCOUNTIF(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B47").getCellFormatText());
	}

	
	protected void testDEVSQ(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("48", Ranges.range(sheet, "B55").getCellFormatText());
	}

	
	protected void testEXPONDIST(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.86", Ranges.range(sheet, "B57").getCellFormatText());
	}

	
	protected void testGAMMAINV(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("10.00", Ranges.range(sheet, "B78").getCellFormatText());
	}

	
	protected void testGAMMALN(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.79", Ranges.range(sheet, "B80").getCellFormatText());
	}

	
	protected void testGEOMEAN(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.48", Ranges.range(sheet, "B82").getCellFormatText());
	}

	
	protected void testHYPGEOMDIST(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.36", Ranges.range(sheet, "B92").getCellFormatText());
	}

	
	protected void testKURT(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("-0.15", Ranges.range(sheet, "B97").getCellFormatText());
	}

	
	protected void testLARGE(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5", Ranges.range(sheet, "B100").getCellFormatText());
	}

	
	protected void testMAX(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("27", Ranges.range(sheet, "B110").getCellFormatText());
	}

	
	protected void testMAXA(Book book) {

		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.5", Ranges.range(sheet, "B112").getCellFormatText());
	}

	
	protected void testMEDIAN(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("3", Ranges.range(sheet, "B114").getCellFormatText());
	}

	
	protected void testMIN(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("2", Ranges.range(sheet, "B116").getCellFormatText());
	}

	
	protected void testMINA(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0", Ranges.range(sheet, "B118").getCellFormatText());
	}

	
	protected void testMODE(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("4", Ranges.range(sheet, "B121").getCellFormatText());
	}

	
	protected void testNORMDIST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.91", Ranges.range(sheet, "B125").getCellFormatText());
	}

	
	protected void testPOISSON(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.12", Ranges.range(sheet, "B142").getCellFormatText());
	}

	
	protected void testRANK(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("3", Ranges.range(sheet, "B149").getCellFormatText());
	}

	
	protected void testSKEW(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.36", Ranges.range(sheet, "B154").getCellFormatText());
	}

	
	protected void testSLOPE(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.31", Ranges.range(sheet, "B156").getCellFormatText());
	}

	
	protected void testSMALL(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("4", Ranges.range(sheet, "B159").getCellFormatText());
	}

	
	protected void testSTDEV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("27.46", Ranges.range(sheet, "B164").getCellFormatText());
	}

	
	protected void testVAR(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("754.27", Ranges.range(sheet, "B189").getCellFormatText());
	}

	
	protected void testVARP(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("678.84", Ranges.range(sheet, "B193").getCellFormatText());
	}

	
	protected void testWEIBULL(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.93", Ranges.range(sheet, "B197").getCellFormatText());
	}

	protected void testIMCONJUGATE(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("3-4i", Ranges.range(sheet, "B45").getCellFormatText());
	}
	
	
	protected void testBESSELI(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.98", Ranges.range(sheet, "B3").getCellFormatText());
	}

	
	protected void testBESSELJ(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.33", Ranges.range(sheet, "B5").getCellFormatText());
	}

	
	protected void testBESSELK(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.28", Ranges.range(sheet, "B7").getCellFormatText());
	}

	
	protected void testBESSELY(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.15", Ranges.range(sheet, "B9").getCellFormatText());
	}

	
	protected void testBIN2DEC(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("100", Ranges.range(sheet, "B11").getCellFormatText());
	}

	
	protected void testBIN2HEX(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("00FB", Ranges.range(sheet, "B13").getCellFormatText());
	}

	
	protected void testBIN2OCT(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B15").getCellFormatText());
	}

	
	protected void testDEC2BIN(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1001", Ranges.range(sheet, "B19").getCellFormatText());
	}

	
	protected void testDEC2HEX(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0064", Ranges.range(sheet, "B21").getCellFormatText());
	}

	
	protected void testDEC2OCT(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("072", Ranges.range(sheet, "B23").getCellFormatText());
	}

	
	protected void testDELTA(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0", Ranges.range(sheet, "B25").getCellFormatText());
	}

	
	protected void testERF(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.71", Ranges.range(sheet, "B27").getCellFormatText());
	}

	
	protected void testERFC(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.16", Ranges.range(sheet, "B29").getCellFormatText());
	}

	
	protected void testGESTEP(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1", Ranges.range(sheet, "B31").getCellFormatText());
	}

	
	protected void testHEX2BIN(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("00001111", Ranges.range(sheet, "B33").getCellFormatText());
	}

	
	protected void testHEX2DEC(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("165", Ranges.range(sheet, "B35").getCellFormatText());
	}

	
	protected void testHEX2OCT(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("017", Ranges.range(sheet, "B37").getCellFormatText());
	}

	
	protected void testIMABS(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("13", Ranges.range(sheet, "B39").getCellFormatText());
	}

	
	protected void testIMAGINARY(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("4", Ranges.range(sheet, "B41").getCellFormatText());
	}

	
	protected void testIMARGUMENT(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.93", Ranges.range(sheet, "B43").getCellFormatText());
	}
	
	protected void testIMDIV(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("5+12i", Ranges.range(sheet, "B49").getCellFormatText());
	}
	
	protected void testIMEXP(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("1.46869393991589+2.28735528717884i", Ranges.range(sheet, "B51").getCellFormatText());
	}

	protected void testIMCOS(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("0.833730025131149-0.988897705762865i", Ranges.range(sheet, "B47").getCellFormatText());
	}

	
	protected void testIMLN(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("1.6094379124341+0.927295218001612i", Ranges.range(sheet, "B53").getCellFormatText());
	}
	
	protected void testIMPRODUCT(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("27+11i", Ranges.range(sheet, "B61").getCellFormatText());
	}

	
	protected void testIMSQRT(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("1.09868411346781+0.455089860562227i", Ranges.range(sheet, "B67").getCellFormatText());

	}

	
	protected void testIMLOG10(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("0.698970004336019+0.402719196273373i", Ranges.range(sheet, "B55").getCellFormatText());

	}

	
	protected void testIMSUM(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("8+i", Ranges.range(sheet, "B71").getCellFormatText());

	}
	
	protected void testIMREAL(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("6", Ranges.range(sheet, "B63").getCellFormatText());
	}

	
	protected void testIMSIN(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("3.85373803791938-27.0168132580039i", Ranges.range(sheet, "B65").getCellFormatText());

	}

	
	protected void testIMSUB(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("8+i", Ranges.range(sheet, "B69").getCellFormatText());

	}

	
	protected void testIMLOG2(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("2.32192809506607+1.33780421255394i", Ranges.range(sheet, "B57").getCellFormatText());

	}

	
	protected void testIMPOWER(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		AssertUtil.assertComplexEquals("-46+9.00000000000001i", Ranges.range(sheet, "B59").getCellFormatText());
	}

	
	protected void testOCT2BIN(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B73").getCellFormatText());
	}

	
	protected void testOCT2DEC(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("44", Ranges.range(sheet, "B75").getCellFormatText());
	}

	
	protected void testOCT2HEX(Book book) {

		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0040", Ranges.range(sheet, "B77").getCellFormatText());
	}

	
	protected void testADDRESS(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("$C$2", Ranges.range(sheet, "B3").getCellFormatText());
	}

	
	protected void testCHOOSE(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2nd", Ranges.range(sheet, "B7").getCellFormatText());
	}

	
	protected void testCOLUMN(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2", Ranges.range(sheet, "B10").getCellFormatText());
	}

	
	protected void testCOLUMNS(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("3", Ranges.range(sheet, "B12").getCellFormatText());
	}

	
	protected void testHLOOKUP(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("4", Ranges.range(sheet, "B14").getCellFormatText());
	}

	
	protected void testHyperLink(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("ZK", Ranges.range(sheet, "B20").getCellFormatText());
	}

	
	protected void testINDIRECT(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("1.333", Ranges.range(sheet, "B26").getCellFormatText());
	}

	
	protected void testLOOKUP(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("orange", Ranges.range(sheet, "B29").getCellFormatText());
	}

	
	protected void testMATCH(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2", Ranges.range(sheet, "B36").getCellFormatText());
	}

	
	protected void testOFFSET(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("offset input", Ranges.range(sheet, "B42").getCellFormatText());
	}

	
	protected void testROW(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("45", Ranges.range(sheet, "B45").getCellFormatText());
	}

	
	protected void testROWS(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("4", Ranges.range(sheet, "B48").getCellFormatText());
	}

	
	protected void testVLOOKUP(Book book) {

		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("2.93", Ranges.range(sheet, "B56").getCellFormatText());
	}
	
	protected void testAREAS(Book book) {
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals("1.00", Ranges.range(sheet, "B5").getCellFormatText());
	}
	
	// 1990/1/1 is Monday, but Excel think it is not work day.
	
	protected void test19900101IsNotWorkDayInExcel(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B32").getCellData().getFormatText()); // =NETWORKDAYS(DATE(1900,1,1),DATE(1900,1,1))
	}
	
	protected void testCOMPLEX(Book book) {
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals(" 3+4i ", Ranges.range(sheet, "B17").getCellFormatText());
	}
	
	protected void testINDEX(Book book) {
		Sheet sheet = book.getSheet("formula-lookup");
		assertEquals(" pears ", Ranges.range(sheet, "B22").getCellFormatText());
		// "accounting format" has space in the front and back.
	}
	
	// expected:<[1]> but was:<[2]>
	// different specification
	protected void testStartDateEmpty(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B29").getCellData().getFormatText()); // =NETWORKDAYS(B30,C30)
		// B30 : blank
		// C30 : 1990/1/2
	}
	
	// This should be check by human
	
	protected void testRANDBETWEEN(Book book) {
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("88", Ranges.range(sheet, "B94").getCellData().getFormatText());
	}
	
	// This should be check by human
	protected void testRAND(Book book) {
		
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("84", Ranges.range(sheet, "B92").getCellData().getFormatText());
	}
	
	protected void testGAMMADIST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.03", Ranges.range(sheet, "B76").getCellFormatText());
	}
	
	protected void testGAMMADISTWithCumulative(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.07", Ranges.range(sheet, "B202").getCellFormatText());
	}
	
	protected void testTDISTWithTwoTail(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.05", Ranges.range(sheet, "B175").getCellFormatText());
	}
	
	protected void testTDISTWithOneTail(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.03", Ranges.range(sheet, "B204").getCellFormatText());
	}

	protected void testFINV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("15.21", Ranges.range(sheet, "B61").getCellFormatText());
	}
	
	protected void testHARMEAN(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("5.03", Ranges.range(sheet, "B89").getCellFormatText());
	}

	protected void testTINV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("1.96", Ranges.range(sheet, "B177").getCellFormatText());
	}
	
	// #NAME?
	protected void testSTDEVP(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("26.05", Ranges.range(sheet, "B168").getCellFormatText());
	}
	
	protected void testHOURWithString(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		Ranges.range(sheet, "B13").setCellEditText("=HOUR(\"15:30\")");
		assertEquals("15", Ranges.range(sheet, "B13").getCellData().getFormatText());
	}
	
	protected void testMINUTEWithString(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		Ranges.range(sheet, "B13").setCellEditText("=MINUTE(\"15:30\")");
		assertEquals("30", Ranges.range(sheet, "B13").getCellData().getFormatText());
	}
	
	protected void testSECONDWithString(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		Ranges.range(sheet, "B13").setCellEditText("=SECOND(\"15:30:55\")");
		assertEquals("55", Ranges.range(sheet, "B13").getCellData().getFormatText());
	}
	
	protected void testVALUE(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1000", Ranges.range(sheet, "B72").getCellFormatText());
	}
	
	// #VALUE!
	protected void testVALUEWithTimeString(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("0.7", Ranges.range(sheet, "B73").getCellFormatText());
	}
	
	// #VALUE!
	protected void testEndDateEmpty(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("-1", Ranges.range(sheet, "B31").getCellData()
				.getFormatText()); // =NETWORKDAYS(C30, B30)
		// B30 : blank
		// C30 : 1990/1/2
	}
	
	protected void testTOUSD2NTD(Book book) {
		Sheet sheet = book.getSheet("formula-custom");
		assertEquals("150", Ranges.range(sheet, "B6").getCellFormatText());
	}
	
	protected void testTOTWD(Book book) {
		Sheet sheet = book.getSheet("formula-custom");
		assertEquals("300", Ranges.range(sheet, "B3").getCellFormatText());
	}

	protected void testLOGEST(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.46", Ranges.range(sheet, "B73").getCellData().getFormatText());	
	}
	
	protected void testMIRR(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("13%", Ranges.range(sheet, "B78").getCellData().getFormatText());	
	}
	
	// #NAME?
	protected void testISPMT(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-64814.81", Ranges.range(sheet, "B70").getCellData().getFormatText());	
	}
	
	protected void testVDB(Book book) {
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("1.32", Ranges.range(sheet, "B117").getCellData().getFormatText());	
	}

	protected void testTIMEVALUE(Book book) {
		Sheet sheet = book.getSheet("formula-datetime");
		assertEquals("0.1", Ranges.range(sheet, "B40").getCellData().getFormatText());
	}
	
	protected void testSERIESSUM(Book book) {
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("0.7071", Ranges.range(sheet, "B104").getCellData().getFormatText());
	}

	// #NAME?
	protected void testCELL(Book book) {
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("48", Ranges.range(sheet, "B48").getCellFormatText());
	}
	
	// #NAME?
	protected void testINFO(Book book) {
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("12.0", Ranges.range(sheet, "B46").getCellFormatText());
	}
	
	// #NAME?
	protected void testASC(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("EXCEL", Ranges.range(sheet, "B2").getCellFormatText());
	}
	
	// #NAME?
	protected void testFINDB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B21").getCellFormatText());
	}
	
	// #NAME?
	protected void testLEFTB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("\u6771", Ranges.range(sheet, "B28").getCellFormatText());
	}
	
	// #NAME?
	protected void testLENB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("11", Ranges.range(sheet, "B32").getCellFormatText());
	}
	
	// #NAME?
	protected void testMIDB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Fluid", Ranges.range(sheet, "B39").getCellFormatText());
	}
	
	// #NAME?
	protected void testSEARCHB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("6", Ranges.range(sheet, "B56").getCellFormatText());
	}
	
	// #NAME?
	protected void testREPLACEB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("abcde*k", Ranges.range(sheet, "B46").getCellFormatText());
	}
	
	// #NAME?
	protected void testRIGHTB(Book book) {
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Price", Ranges.range(sheet, "B52").getCellFormatText());
	}
	
	// #NAME?
	protected void testFORECAST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("10.61", Ranges.range(sheet, "B67").getCellFormatText());
	}
	
	// #NAME?
	protected void testNORMSDIST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.91", Ranges.range(sheet, "B129").getCellFormatText());
	}
	
	// #NAME?
	protected void testCRITBINOM(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("4", Ranges.range(sheet, "B53").getCellFormatText());
	}
	
	// #NAME?
	protected void testINTERCEPT(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.05", Ranges.range(sheet, "B94").getCellFormatText());
	}
	
	// #NAME?
	protected void testFISHERINV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.75", Ranges.range(sheet, "B65").getCellFormatText());
	}
	
	// #NAME?
	protected void testSTDEVA(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("27.46", Ranges.range(sheet, "B166").getCellFormatText());
	}
	
	// #NAME?
	protected void testPERMUT(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("970200", Ranges.range(sheet, "B140").getCellFormatText());
	}
	
	// #NAME?
	protected void testLOGINV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("4.00", Ranges.range(sheet, "B106").getCellFormatText());
	}
	
	// #NAME?
	protected void testLINEST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("2", Ranges.range(sheet, "B103").getCellFormatText());
	}
	
	// #NAME?
	protected void testFREQUENCY(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("1", Ranges.range(sheet, "B70").getCellFormatText());
	}
	
	// #NAME?
	protected void testGROWTH(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("320196.72", Ranges.range(sheet, "B85").getCellFormatText());
	}
	
	// #NAME?
	protected void testFISHER(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.97", Ranges.range(sheet, "B63").getCellFormatText());
	}
	
	// #NAME?
	protected void testPERCENTRANK(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.33", Ranges.range(sheet, "B138").getCellFormatText());
	}
	
	// #NAME?
	protected void testCORREL(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.9971", Ranges.range(sheet, "B37").getCellFormatText());
	}
	
	// #NAME?
	protected void testPERCENTILE(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("1.9", Ranges.range(sheet, "B136").getCellFormatText());
	}
	
	// #NAME?
	
	protected void testSTDEVPA(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("26.05", Ranges.range(sheet, "B170").getCellFormatText());
	}
	
	// #NAME?
	protected void testAVERAGEIF(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("14000", Ranges.range(sheet, "B11").getCellFormatText());
	}
	
	// #NAME?
	protected void testQUARTILE(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("3.5", Ranges.range(sheet, "B147").getCellFormatText());
	}
	
	// #NAME?
	protected void testMORMINV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("42.00", Ranges.range(sheet, "B127").getCellFormatText());
	}
	
	// #NAME?
	protected void testNEGBINOMDIST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.06", Ranges.range(sheet, "B123").getCellFormatText());
	}
	
	// #NAME?
	protected void testCHITEST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.000308", Ranges.range(sheet, "B27").getCellFormatText());
	}
	
	// #NAME?
	protected void testBETADIST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.69", Ranges.range(sheet, "B17").getCellFormatText());
	}
	
	// #NAME?
	protected void testVARA(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("754.27", Ranges.range(sheet, "B191").getCellFormatText());
	}
	
	// #NAME?
	protected void testPROB(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.1", Ranges.range(sheet, "B144").getCellFormatText());
	}
	
	// #NAME?
	protected void testZTEST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.09", Ranges.range(sheet, "B199").getCellFormatText());
	}
	
	// #NAME?
	protected void testVARPA(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("678.84", Ranges.range(sheet, "B195").getCellFormatText());
	}
	
	// #NAME?
	protected void testTTEST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.20", Ranges.range(sheet, "B186").getCellFormatText());
	}
	
	// #NAME?
	protected void testTREND(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("133953.33", Ranges.range(sheet, "B179").getCellFormatText());
	}
	
	// #NAME?
	protected void testSTEYX(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("3.31", Ranges.range(sheet, "B172").getCellFormatText());
	}
	
	// #NAME?
	protected void testFTEST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.65", Ranges.range(sheet, "B73").getCellFormatText());
	}
	
	protected void testFDIST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.01", Ranges.range(sheet, "B59").getCellFormatText());
	}
	
	// #NAME?
	protected void testCOVAR(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("5.2", Ranges.range(sheet, "B50").getCellFormatText());
	}
	
	// #NAME?
	protected void testLOGNORMDIST(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.04", Ranges.range(sheet, "B108").getCellFormatText());
	}
	
	// #NAME?
	protected void testNORMSINV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("1.33", Ranges.range(sheet, "B131").getCellFormatText());
	}
	
	// #NAME?
	protected void testSTANDARDIZE(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("1.33", Ranges.range(sheet, "B162").getCellFormatText());
	}
	
	// #NAME?
	protected void testTRIMMEAN(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("3.78", Ranges.range(sheet, "B183").getCellFormatText());
	}
	
	// #NAME?
	protected void testCONFIDENCE(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.69", Ranges.range(sheet, "B35").getCellFormatText());
	}
	
	// #NAME?
	protected void testRSQ(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.061", Ranges.range(sheet, "B151").getCellFormatText());
	}
	
	// #NAME?
	protected void testBETAINV(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("2", Ranges.range(sheet, "B19").getCellFormatText());
	}
	
	// #NAME?
	protected void testAVERAGEIFS(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("87.5", Ranges.range(sheet, "B14").getCellFormatText());
	}
	
	// #NAME?
	protected void testPEARSON(Book book) {
		Sheet sheet = book.getSheet("formula-statistical");
		assertEquals("0.70", Ranges.range(sheet, "B133").getCellFormatText());
	}

}
