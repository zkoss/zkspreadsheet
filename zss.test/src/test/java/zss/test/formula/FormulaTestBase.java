package zss.test.formula;

import static org.junit.Assert.assertEquals;

import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;
import zss.test.formula.AssertUtil;

/**
 * all formula assert utility.
 * 
 * @author kuro
 */
public class FormulaTestBase {
	
	// spreadsheet doesn't change format into currency automatically.
	protected void testDOLLAR(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("$1,234.57", Ranges.range(sheet, "B13").getFormatText().getCellFormatResult().text);
	}

	protected void testACCRINT(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("16.67", Ranges.range(sheet, "B3").getFormatText().getCellFormatResult().text);
	}

	
	protected void testACCRINTM(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("20.55", Ranges.range(sheet, "B6").getFormatText().getCellFormatResult().text);
	}

	
	protected void testAMORDEGRC(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("776", Ranges.range(sheet, "B9").getFormatText().getCellFormatResult().text);
	}

	
	protected void testAMORLINC(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("360", Ranges.range(sheet, "B12").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUPDAYBS(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("71", Ranges.range(sheet, "B15").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUPDAYS(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("181", Ranges.range(sheet, "B18").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUPDAYSNC(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("110", Ranges.range(sheet, "B21").getFormatText().getCellFormatResult().text);
	}
	
	protected void testCOUPNCD(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("2011/5/15", Ranges.range(sheet, "B24").getFormatText().getCellFormatResult().text);
	}

	protected void testCOUPPCD(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("2006/11/15", Ranges.range(sheet, "B30").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUPNUM(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("4", Ranges.range(sheet, "B27").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCUMIPMT(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-11135.23", Ranges.range(sheet, "B33").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCUMPRINC(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-934.11", Ranges.range(sheet, "B36").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("186083.33", Ranges.range(sheet, "B39").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDDB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1.32", Ranges.range(sheet, "B42").getFormatText().getCellFormatResult().text);
	}

	protected void testDISC(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B45").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDOLLARDE(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1.125", Ranges.range(sheet, "B47").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDOLLARFR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1.02", Ranges.range(sheet, "B49").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDURATION(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("5.99", Ranges.range(sheet, "B51").getFormatText().getCellFormatResult().text);
	}

	
	protected void testEFFECT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B54").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("2581.40", Ranges.range(sheet, "B57").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFVSCHEDULE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1.33", Ranges.range(sheet, "B59").getFormatText().getCellFormatResult().text);
	}

	
	protected void testINTRATE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B62").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIPMT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-292.45", Ranges.range(sheet, "B64").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIRR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-2.1%", Ranges.range(sheet, "B67").getFormatText().getCellFormatResult().text);
	}

	
	protected void testNOMINAL(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B81").getFormatText().getCellFormatResult().text);
	}

	
	protected void testNPER(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("59.67", Ranges.range(sheet, "B84").getFormatText().getCellFormatResult().text);
	}

	
	protected void testNPV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("41922.06", Ranges.range(sheet, "B87").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPMT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-1037.03", Ranges.range(sheet, "B89").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPPMT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-75.62", Ranges.range(sheet, "B91").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPRICE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("94.63", Ranges.range(sheet, "B93").getFormatText().getCellFormatResult().text);

	}

	
	protected void testPRICEDISC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("99.80", Ranges.range(sheet, "B95").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPRICEMAT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("99.98", Ranges.range(sheet, "B97").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-59777.15", Ranges.range(sheet, "B99").getFormatText().getCellFormatResult().text);
	}

	
	protected void testRATE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1%", Ranges.range(sheet, "B102").getFormatText().getCellFormatResult().text);
	}

	
	protected void testRECEIVED(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1014584.65", Ranges.range(sheet, "B104").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSLN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("2250", Ranges.range(sheet, "B106").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSYD(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("4090.91", Ranges.range(sheet, "B108").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTBILLEQ(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.09", Ranges.range(sheet, "B110").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTBILLPRICE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("98.45", Ranges.range(sheet, "B112").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTBILLYIELD(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.09", Ranges.range(sheet, "B115").getFormatText().getCellFormatResult().text);
	}

	
	protected void testXNPV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("2086.65", Ranges.range(sheet, "B119").getFormatText().getCellFormatResult().text);
	}

	
	protected void testYIELD(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B123").getFormatText().getCellFormatResult().text);
	}

	
	protected void testYIELDDISC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.05", Ranges.range(sheet, "B125").getFormatText().getCellFormatResult().text);
	}

	
	protected void testYIELDMAT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("0.06", Ranges.range(sheet, "B127").getFormatText().getCellFormatResult().text);
	}

	
	protected void testAND(Book book) {

		Worksheet sheet = book.getWorksheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B4").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFALSE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-logical");
		assertEquals("FALSE", Ranges.range(sheet, "B6").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIF(Book book) {

		Worksheet sheet = book.getWorksheet("formula-logical");
		assertEquals("over 10", Ranges.range(sheet, "B8").getFormatText().getRichTextString().getString());
	}

	
	protected void testIFERROR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-logical");
		assertEquals("6", Ranges.range(sheet, "B11").getFormatText().getCellFormatResult().text);
	}

	
	protected void testNOT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B14").getFormatText().getCellFormatResult().text);
	}

	
	protected void testOR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-logical");
		assertEquals("FALSE", Ranges.range(sheet, "B16").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTRUE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-logical");
		assertEquals("TRUE", Ranges.range(sheet, "B18").getFormatText().getCellFormatResult().text);
	}

	
	protected void testABS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B4").getFormatText().getCellFormatResult().text);
	}

	
	protected void testACOS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2.09", Ranges.range(sheet, "B7").getFormatText().getCellFormatResult().text);
	}

	
	protected void testACOSH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0", Ranges.range(sheet, "B9").getFormatText().getCellFormatResult().text);
	}

	
	protected void testASIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("-0.5236", Ranges.range(sheet, "B11").getFormatText().getCellFormatResult().text);
	}

	
	protected void testASINH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("-1.65", Ranges.range(sheet, "B13").getFormatText().getCellFormatResult().text);
	}

	
	protected void testATAN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0.79", Ranges.range(sheet, "B15").getFormatText().getCellFormatResult().text);
	}

	
	protected void testATAN2(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0.79", Ranges.range(sheet, "B17").getFormatText().getCellFormatResult().text);
	}

	
	protected void testATANH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1.00", Ranges.range(sheet, "B19").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCEILING(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B21").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOMBIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("28", Ranges.range(sheet, "B23").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0.50", Ranges.range(sheet, "B25").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOSH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("27.31", Ranges.range(sheet, "B27").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDEGREES(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("180", Ranges.range(sheet, "B29").getFormatText().getCellFormatResult().text);
	}

	
	protected void testEVEN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B31").getFormatText().getCellFormatResult().text);
	}

	
	protected void testEXP(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2.72", Ranges.range(sheet, "B33").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFACT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("120", Ranges.range(sheet, "B35").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFACTDOUBLE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("48", Ranges.range(sheet, "B37").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFLOOR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B39").getFormatText().getCellFormatResult().text);
	}

	
	protected void testGCD(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B41").getFormatText().getCellFormatResult().text);
	}

	
	protected void testINT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("8", Ranges.range(sheet, "B43").getFormatText().getCellFormatResult().text);
	}

	
	protected void testLCM(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("10", Ranges.range(sheet, "B45").getFormatText().getCellFormatResult().text);
	}

	
	protected void testLN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("4.4543", Ranges.range(sheet, "B47").getFormatText().getCellFormatResult().text);
	}

	
	protected void testLOG(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B49").getFormatText().getCellFormatResult().text);
	}

	
	protected void testLOG10(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B51").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMDETERM(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("88", Ranges.range(sheet, "B53").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMINVERSE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0", Ranges.range(sheet, "B59").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMMULT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B63").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMOD(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B71").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMROUND(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("9", Ranges.range(sheet, "B73").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMULTINOMIA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1260", Ranges.range(sheet, "B75").getFormatText().getCellFormatResult().text);
	}

	
	protected void testODD(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B77").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPI(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("3.1416", Ranges.range(sheet, "B79").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPOWER(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("25", Ranges.range(sheet, "B81").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPRODUCT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2250", Ranges.range(sheet, "B83").getFormatText().getCellFormatResult().text);
	}

	
	protected void testQUOTIENT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2", Ranges.range(sheet, "B88").getFormatText().getCellFormatResult().text);
	}

	
	protected void testRADIANS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("4.7124", Ranges.range(sheet, "B90").getFormatText().getCellFormatResult().text);
	}

	
	protected void testROMAN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("CDXCIX", Ranges.range(sheet, "B96").getFormatText().getRichTextString().getString());
	}

	
	protected void testROUND(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("2.2", Ranges.range(sheet, "B98").getFormatText().getCellFormatResult().text);
	}

	
	protected void testROUNDDOWN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("3", Ranges.range(sheet, "B100").getFormatText().getCellFormatResult().text);
	}

	
	protected void testROUNDUP(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B102").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSIGN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1", Ranges.range(sheet, "B111").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0.00", Ranges.range(sheet, "B113").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSINH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1.1752", Ranges.range(sheet, "B115").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSQRT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("4", Ranges.range(sheet, "B117").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSQRTPI(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("1.7725", Ranges.range(sheet, "B119").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUBTOTAL(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("303", Ranges.range(sheet, "B121").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUM(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("5", Ranges.range(sheet, "B127").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUMIF(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("63000", Ranges.range(sheet, "B129").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUMIFS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("20", Ranges.range(sheet, "B135").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUMPRODUCT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("156", Ranges.range(sheet, "B145").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUMSQ(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("25", Ranges.range(sheet, "B150").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUMX2MY2(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("-55", Ranges.range(sheet, "B153").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUMX2PY2(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("521", Ranges.range(sheet, "B162").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUMXMY2(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("79", Ranges.range(sheet, "B171").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTAN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0.9992", Ranges.range(sheet, "B180").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTANH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("-0.96", Ranges.range(sheet, "B182").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTRUNC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("8", Ranges.range(sheet, "B184").getFormatText().getCellFormatResult().text);
	}

	// FIXME
//	protected void testDATE(Book book) {
//
//		Worksheet sheet = book.getWorksheet("formula-datetime");
//		CellOperationUtil.applyDataFormat(Ranges.range(sheet, "B3"), "Y/M/D");
//		assertEquals("08/1/1", Ranges.range(sheet, "B3").getFormatText().getCellFormatResult().text);
//		// '08/01/01' is correct and fixed in zpoi 3.9.
//	}

	
	protected void testDATEVALUE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("39448", Ranges.range(sheet, "B6").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDAY(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("15", Ranges.range(sheet, "B8").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDAY360(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B10").getFormatText().getCellFormatResult().text);
	}

	
	protected void testHOUR(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("3", Ranges.range(sheet, "B13").getFormatText().getCellFormatResult().text);
	}
	
	protected void testMINUTE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("48", Ranges.range(sheet, "B15").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMONTH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("4", Ranges.range(sheet, "B17").getFormatText().getCellFormatResult().text);
	}

	
	protected void testNETWORKDAYS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("22", Ranges.range(sheet, "B19").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2013,4,1),DATE(2013,4,30))
		assertEquals("108", Ranges.range(sheet, "B20").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2008,10,1),
																						// DATE(2009,3,1))
	}

	
	protected void testNetworkDaysStartDateIsHoliday(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("5", Ranges.range(sheet, "B21").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,9))
	}

	
	protected void testNetworkDaysAllHoliday(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B22").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,1))
		assertEquals("0", Ranges.range(sheet, "B23").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,2))
	}

	
	protected void testNetworkDaysSpecificHoliday(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("107", Ranges.range(sheet, "B24").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2008,10,1),
																						// DATE(2009,3,1),
																						// DATE(2008,11,26))
		assertEquals("1", Ranges.range(sheet, "B25").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2013,6,1),
																						// DATE(2013,6,4),
																						// DATE(2013,6,3))
	}

	
	protected void testStartDateEqualsEndDate(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B26").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2013,6,28),
																						// DATE(2013,6,28))
	}

	
	protected void testStartDateLaterThanEndDate(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("#VALUE!", Ranges.range(sheet, "B27").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(2013,6,2),
																							// DATE(2013,6,1))
	}

	
	protected void testEmptyDate(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B28").getFormatText().getCellFormatResult().text); // =NETWORKDAYS(E28,F28)
		// E28 and F28 are blank
	}

	
	protected void testSECOND(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("18", Ranges.range(sheet, "B36").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTIME(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("0.5", Ranges.range(sheet, "B38").getFormatText().getCellFormatResult().text);
	}

	
	protected void testWEEKDAY(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("5", Ranges.range(sheet, "B44").getFormatText().getCellFormatResult().text);
	}
	
	protected void testWORKDAY(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("2013/4/5", Ranges.range(sheet, "B46")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,1),4)
		assertEquals("2009/4/30", Ranges.range(sheet, "B47")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2008,10,1),151)
	}

	protected void testStartDateIsHoliday(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("2013/6/3", Ranges.range(sheet, "B48")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,6,1),1)
		assertEquals("2013/5/31", Ranges.range(sheet, "B49")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,6,1), -1)
	}

	protected void testEndDateIsHoliday(Book book)  {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("2013/4/8", Ranges.range(sheet, "B50")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,1),5)
		assertEquals("2013/4/8", Ranges.range(sheet, "B51")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,5),1)
	}

	protected void testNegativeWorkdayEndDateIsHolday(Book book)  {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("2013/3/29", Ranges.range(sheet, "B52")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,1),-1)
		assertEquals("2013/5/31", Ranges.range(sheet, "B53")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,6,7),-5)
	}

	protected void testWorkdayBoundary(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("2013/4/1", Ranges.range(sheet, "B54")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,1),0)
		assertEquals("2013/6/1", Ranges.range(sheet, "B55")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,6,1),0)
	}

	protected void testWorkdaySpecifiedholiday(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("2013/4/5", Ranges.range(sheet, "B56")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,1),3, DATE(2013,4,2))
		assertEquals("2013/4/3", Ranges.range(sheet, "B57")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,7),
									// -3,DATE(2013,4,2))
		assertEquals("2013/4/1", Ranges.range(sheet, "B58")
				.getFormatText().getCellFormatResult().text); // =WORKDAY(DATE(2013,4,3),
									// -1,DATE(2013,4,2))
	}

	
	protected void testYEAR(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("2008", Ranges.range(sheet, "B61").getFormatText().getCellFormatResult().text); // =YEAR(DATE(2008,1,1))
	}

	
	protected void testYEARFRAC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals(0.58, (Double)Ranges.range(sheet, "B63").getValue(), 0.005); // =YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))
	}

	
	protected void testCHAR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("A", Ranges.range(sheet, "B4").getFormatText().getRichTextString().getString());
	}

	
	protected void testLOWER(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("e. e. cummings", Ranges.range(sheet, "B34").getFormatText().getRichTextString().getString());
	}

	
	protected void testCODE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("65", Ranges.range(sheet, "B9").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCLEAN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("text", Ranges.range(sheet, "B6").getFormatText().getRichTextString().getString());
	}

	
	protected void testCONCATENATE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("ZK", Ranges.range(sheet, "B11").getFormatText().getRichTextString().getString());
	}

	
	protected void testEXACT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("TRUE", Ranges.range(sheet, "B15").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFIND(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B19").getFormatText().getCellFormatResult().text);
	}

	
	protected void testFIXED(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("1,234.6", Ranges.range(sheet, "B23").getFormatText().getCellFormatResult().text);
	}

	
	protected void testLEFT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("\u6771 \u4EAC ", Ranges.range(sheet, "B26").getFormatText().getRichTextString().getString());
	}

	
	protected void testLEN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("11", Ranges.range(sheet, "B30").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMID(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("Fluid", Ranges.range(sheet, "B37").getFormatText().getRichTextString().getString());
	}

	
	protected void testPROPER(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("This Is A Title", Ranges.range(sheet, "B41").getFormatText().getRichTextString().getString());
	}

	
	protected void testREPLACE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("abcde*k", Ranges.range(sheet, "B44").getFormatText().getRichTextString().getString());
	}

	
	protected void testREPT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("*-*-*-", Ranges.range(sheet, "B48").getFormatText().getRichTextString().getString());
	}

	
	protected void testRIGHT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("Price", Ranges.range(sheet, "B50").getFormatText().getRichTextString().getString());
	}

	
	protected void testSEARCH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("6", Ranges.range(sheet, "B54").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSUBSTITUTE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("Quarter 2, 2008", Ranges.range(sheet, "B58").getFormatText().getRichTextString().getString());
	}

	
	protected void testT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("Sale Price", Ranges.range(sheet, "B62").getFormatText().getRichTextString().getString());
	}

	
	protected void testTEXT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("Date: 2007-08-06", Ranges.range(sheet, "B64").getFormatText().getRichTextString().getString());
	}

	
	protected void testTRIM(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("revenue in quarter 1", Ranges.range(sheet, "B67").getFormatText().getRichTextString().getString());
	}

	
	protected void testUPPER(Book book) {

		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("TOTAL", Ranges.range(sheet, "B69").getFormatText().getRichTextString().getString());
	}

	
	protected void testAVEDEV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("1.02", Ranges.range(sheet, "B3").getFormatText().getCellFormatResult().text);
	}

	
	protected void testAVERAGE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("11", Ranges.range(sheet, "B6").getFormatText().getCellFormatResult().text);
	}

	
	protected void testERRORTYPE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("1", Ranges.range(sheet, "B3").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISBLANK(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B10").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISERR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B12").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISERROR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B15").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISEVEN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B18").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISLOGICAL(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B21").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISNA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B23").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISNONTEXT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B26").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISNUMBER(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B29").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISODD(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B31").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISREF(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("FALSE", Ranges.range(sheet, "B33").getFormatText().getCellFormatResult().text);
	}

	
	protected void testISTEXT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("TRUE", Ranges.range(sheet, "B35").getFormatText().getCellFormatResult().text);
	}

	
	protected void testN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("7", Ranges.range(sheet, "B38").getFormatText().getCellFormatResult().text);
	}

	
	protected void testNA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("#N/A", Ranges.range(sheet, "B41").getFormatText().getCellFormatResult().text);
	}

	
	protected void testTYPE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("2", Ranges.range(sheet, "B43").getFormatText().getCellFormatResult().text);
	}

	
	protected void testAVERAGEA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("5.6", Ranges.range(sheet, "B8").getFormatText().getCellFormatResult().text);
	}

	
	protected void testBIOMDIST(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.21", Ranges.range(sheet, "B21").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCHIDIST(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.05", Ranges.range(sheet, "B23").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCHIINV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("18.31", Ranges.range(sheet, "B25").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUNT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B41").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUNTA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("6", Ranges.range(sheet, "B43").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUNTBLANK(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B45").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOUNTIF(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B47").getFormatText().getCellFormatResult().text);
	}

	
	protected void testDEVSQ(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("48", Ranges.range(sheet, "B55").getFormatText().getCellFormatResult().text);
	}

	
	protected void testEXPONDIST(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.86", Ranges.range(sheet, "B57").getFormatText().getCellFormatResult().text);
	}

	
	protected void testGAMMAINV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("10.00", Ranges.range(sheet, "B78").getFormatText().getCellFormatResult().text);
	}

	
	protected void testGAMMALN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("1.79", Ranges.range(sheet, "B80").getFormatText().getCellFormatResult().text);
	}

	
	protected void testGEOMEAN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("5.48", Ranges.range(sheet, "B82").getFormatText().getCellFormatResult().text);
	}

	
	protected void testHYPGEOMDIST(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.36", Ranges.range(sheet, "B92").getFormatText().getCellFormatResult().text);
	}

	
	protected void testKURT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("-0.15", Ranges.range(sheet, "B97").getFormatText().getCellFormatResult().text);
	}

	
	protected void testLARGE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("5", Ranges.range(sheet, "B100").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMAX(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("27", Ranges.range(sheet, "B110").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMAXA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.5", Ranges.range(sheet, "B112").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMEDIAN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B114").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B116").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMINA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0", Ranges.range(sheet, "B118").getFormatText().getCellFormatResult().text);
	}

	
	protected void testMODE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B121").getFormatText().getCellFormatResult().text);
	}

	
	protected void testNORMDIST(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.91", Ranges.range(sheet, "B125").getFormatText().getCellFormatResult().text);
	}

	
	protected void testPOISSON(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.12", Ranges.range(sheet, "B142").getFormatText().getCellFormatResult().text);
	}

	
	protected void testRANK(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("3", Ranges.range(sheet, "B149").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSKEW(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.36", Ranges.range(sheet, "B154").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSLOPE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.31", Ranges.range(sheet, "B156").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSMALL(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B159").getFormatText().getCellFormatResult().text);
	}

	
	protected void testSTDEV(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("27.46", Ranges.range(sheet, "B164").getFormatText().getCellFormatResult().text);
	}

	
	protected void testVAR(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("754.27", Ranges.range(sheet, "B189").getFormatText().getCellFormatResult().text);
	}

	
	protected void testVARP(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("678.84", Ranges.range(sheet, "B193").getFormatText().getCellFormatResult().text);
	}

	
	protected void testWEIBULL(Book book) {

		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.93", Ranges.range(sheet, "B197").getFormatText().getCellFormatResult().text);
	}

	
	protected void testBESSELI(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0.98", Ranges.range(sheet, "B3").getFormatText().getCellFormatResult().text);
	}

	
	protected void testBESSELJ(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0.33", Ranges.range(sheet, "B5").getFormatText().getCellFormatResult().text);
	}

	
	protected void testBESSELK(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0.28", Ranges.range(sheet, "B7").getFormatText().getCellFormatResult().text);
	}

	
	protected void testBESSELY(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0.15", Ranges.range(sheet, "B9").getFormatText().getCellFormatResult().text);
	}

	
	protected void testBIN2DEC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("100", Ranges.range(sheet, "B11").getFormatText().getCellFormatResult().text);
	}

	
	protected void testBIN2HEX(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("00FB", Ranges.range(sheet, "B13").getFormatText().getRichTextString().getString());
	}

	
	protected void testBIN2OCT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B15").getFormatText().getRichTextString().getString());
	}

	
	protected void testDEC2BIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("1001", Ranges.range(sheet, "B19").getFormatText().getRichTextString().getString());
	}

	
	protected void testDEC2HEX(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0064", Ranges.range(sheet, "B21").getFormatText().getRichTextString().getString());
	}

	
	protected void testDEC2OCT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("072", Ranges.range(sheet, "B23").getFormatText().getRichTextString().getString());
	}

	
	protected void testDELTA(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0", Ranges.range(sheet, "B25").getFormatText().getCellFormatResult().text);
	}

	
	protected void testERF(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0.71", Ranges.range(sheet, "B27").getFormatText().getCellFormatResult().text);
	}

	
	protected void testERFC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0.16", Ranges.range(sheet, "B29").getFormatText().getCellFormatResult().text);
	}

	
	protected void testGESTEP(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("1", Ranges.range(sheet, "B31").getFormatText().getCellFormatResult().text);
	}

	
	protected void testHEX2BIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("00001111", Ranges.range(sheet, "B33").getFormatText().getRichTextString().getString());
	}

	
	protected void testHEX2DEC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("165", Ranges.range(sheet, "B35").getFormatText().getCellFormatResult().text);
	}

	
	protected void testHEX2OCT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("017", Ranges.range(sheet, "B37").getFormatText().getRichTextString().getString());
	}

	
	protected void testIMABS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("13", Ranges.range(sheet, "B39").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIMAGINARY(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("4", Ranges.range(sheet, "B41").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIMARGUMENT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0.93", Ranges.range(sheet, "B43").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIMCOS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("0.833730025131149-0.988897705762865i", Ranges.range(sheet, "B47").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIMLN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("1.6094379124341+0.927295218001612i", Ranges.range(sheet, "B53").getFormatText().getRichTextString().getString());

	}

	
	protected void testIMSQRT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("1.09868411346781+0.455089860562227i", Ranges.range(sheet, "B67").getFormatText().getRichTextString().getString());

	}

	
	protected void testIMLOG10(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("0.698970004336019+0.402719196273373i", Ranges.range(sheet, "B55").getFormatText().getRichTextString().getString());

	}

	
	protected void testIMSUM(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("8+i", Ranges.range(sheet, "B71").getFormatText().getRichTextString().getString());

	}
	
	protected void testIMREAL(Book book) {
		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("6", Ranges.range(sheet, "B63").getFormatText().getCellFormatResult().text);
	}

	
	protected void testIMSIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("3.85373803791938-27.0168132580039i", Ranges.range(sheet, "B65").getFormatText().getRichTextString().getString());

	}

	
	protected void testIMSUB(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("8+i", Ranges.range(sheet, "B69").getFormatText().getRichTextString().getString());

	}

	
	protected void testIMLOG2(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("2.32192809506607+1.33780421255394i", Ranges.range(sheet, "B57").getFormatText().getRichTextString().getString());

	}

	
	protected void testIMPOWER(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		AssertUtil.assertComplexEquals("-46+9.00000000000001i", Ranges.range(sheet, "B59").getFormatText().getRichTextString().getString());
	}

	
	protected void testOCT2BIN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("011", Ranges.range(sheet, "B73").getFormatText().getRichTextString().getString());
	}

	
	protected void testOCT2DEC(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("44", Ranges.range(sheet, "B75").getFormatText().getCellFormatResult().text);
	}

	
	protected void testOCT2HEX(Book book) {

		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals("0040", Ranges.range(sheet, "B77").getFormatText().getRichTextString().getString());
	}

	
	protected void testADDRESS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("$C$2", Ranges.range(sheet, "B3").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCHOOSE(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("2nd", Ranges.range(sheet, "B7").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOLUMN(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("2", Ranges.range(sheet, "B10").getFormatText().getCellFormatResult().text);
	}

	
	protected void testCOLUMNS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("3", Ranges.range(sheet, "B12").getFormatText().getCellFormatResult().text);
	}

	
	protected void testHLOOKUP(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("4", Ranges.range(sheet, "B14").getFormatText().getCellFormatResult().text);
	}

	
	protected void testHyperLink(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("ZK", Ranges.range(sheet, "B20").getFormatText().getRichTextString().getString());
	}

	
	protected void testINDIRECT(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("1.333", Ranges.range(sheet, "B26").getFormatText().getCellFormatResult().text);
	}

	
	protected void testLOOKUP(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("orange", Ranges.range(sheet, "B29").getFormatText().getRichTextString().getString());
	}

	
	protected void testMATCH(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("2", Ranges.range(sheet, "B36").getFormatText().getCellFormatResult().text);
	}

	
	protected void testOFFSET(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("offset input", Ranges.range(sheet, "B42").getFormatText().getRichTextString().getString());
	}

	
	protected void testROW(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("45", Ranges.range(sheet, "B45").getFormatText().getCellFormatResult().text);
	}

	
	protected void testROWS(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("4", Ranges.range(sheet, "B48").getFormatText().getCellFormatResult().text);
	}

	
	protected void testVLOOKUP(Book book) {

		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("2.93", Ranges.range(sheet, "B56").getFormatText().getCellFormatResult().text);
	}
	
	protected void testAREAS(Book book) {
		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals("1.00", Ranges.range(sheet, "B5").getFormatText().getCellFormatResult().text);
	}
	
	// 1990/1/1 is Monday, but Excel think it is not work day.
	
	protected void test19900101IsNotWorkDayInExcel(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("0", Ranges.range(sheet, "B32")
				.getFormatText().getCellFormatResult().text); // =NETWORKDAYS(DATE(1900,1,1),DATE(1900,1,1))
	}
	
	protected void testCOMPLEX(Book book) {
		Worksheet sheet = book.getWorksheet("formula-engineering");
		assertEquals(" 3+4i ", Ranges.range(sheet, "B17").getFormatText().getCellFormatResult().text);
	}
	
	protected void testINDEX(Book book) {
		Worksheet sheet = book.getWorksheet("formula-lookup");
		assertEquals(" pears ", Ranges.range(sheet, "B22").getFormatText().getCellFormatResult().text);
		// "accounting format" has space in the front and back.
	}
	
	// expected:<[1]> but was:<[2]>
	// different specification
	protected void testStartDateEmpty(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("1", Ranges.range(sheet, "B29")
				.getFormatText().getCellFormatResult().text); // =NETWORKDAYS(B30,C30)
		// B30 : blank
		// C30 : 1990/1/2
	}
	
	// This should be check by human
	
	protected void testRANDBETWEEN(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("88", Ranges.range(sheet, "B94").getFormatText().getCellFormatResult().text);
	}
	
	// This should be check by human
	
	protected void testRAND(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("84", Ranges.range(sheet, "B92").getFormatText().getCellFormatResult().text);
	}
	
	// expected:<[0.03]> but was:<[-1.00]>
	protected void testGAMMADIST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.03", Ranges.range(sheet, "B76").getFormatText().getCellFormatResult().text);
	}
	
	protected void testGAMMADISTWithCumulative(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.07", Ranges.range(sheet, "B202").getFormatText().getCellFormatResult().text);
	}
	
	protected void testTDISTWithTwoTail(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.05", Ranges.range(sheet, "B175").getFormatText().getCellFormatResult().text);
	}
	
	protected void testTDISTWithOneTail(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.03", Ranges.range(sheet, "B204").getFormatText().getCellFormatResult().text);
	}
	
	// expected:<[15.2]1> but was:<[0.1]1>
	
	protected void testFINV(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("15.21", Ranges.range(sheet, "B61").getFormatText().getCellFormatResult().text);
	}
	
	// expected:<[5.03]> but was:<[0.39]>
	
	protected void testHARMEAN(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("5.03", Ranges.range(sheet, "B89").getFormatText().getCellFormatResult().text);
	}
	
	// expected:<1.[96]> but was:<1.[63]>
	
	protected void testTINV(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("1.96", Ranges.range(sheet, "B177").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testSTDEVP(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("26.05", Ranges.range(sheet, "B168").getFormatText().getCellFormatResult().text);
	}
	
//	// #VALUE!
//	protected void testHOURWithString(Book book) {
//		Worksheet sheet = book.getWorksheet("formula-datetime");
//		Ranges.range(sheet, "B13").setCellEditText("=HOUR(\"15:30\")");
//		assertEquals("3", Ranges.range(sheet, "B13").getFormatText().getCellFormatResult().text);
//	}
	
	protected void testVALUE(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("1000", Ranges.range(sheet, "B72").getFormatText().getCellFormatResult().text);
	}
	
	// #VALUE!
	protected void testVALUEWithTimeString(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("0.7", Ranges.range(sheet, "B73").getFormatText().getCellFormatResult().text);
	}
	
	// #VALUE!
	protected void testEndDateEmpty(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("-1", Ranges.range(sheet, "B31")
				.getFormatText().getCellFormatResult().text); // =NETWORKDAYS(C30, B30)
		// B30 : blank
		// C30 : 1990/1/2
	}
	
	protected void testTOUSD2NTD(Book book) {
		Worksheet sheet = book.getWorksheet("formula-custom");
		assertEquals("150", Ranges.range(sheet, "B6").getFormatText().getCellFormatResult().text);
	}
	
	protected void testTOTWD(Book book) {
		Worksheet sheet = book.getWorksheet("formula-custom");
		assertEquals("300", Ranges.range(sheet, "B3").getFormatText().getCellFormatResult().text);
	}

	protected void testLOGEST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1.46", Ranges.range(sheet, "B73")
				.getFormatText().getCellFormatResult().text);	
	}
	
	protected void testMIRR(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("13%", Ranges.range(sheet, "B78")
				.getFormatText().getCellFormatResult().text);	
	}
	
	// #NAME?
	
	protected void testISPMT(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("-64814.81", Ranges.range(sheet, "B70")
				.getFormatText().getCellFormatResult().text);	
	}
	
	protected void testVDB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-financial");
		assertEquals("1.32", Ranges.range(sheet, "B117")
				.getFormatText().getCellFormatResult().text);	
	}

	protected void testTIMEVALUE(Book book) {
		Worksheet sheet = book.getWorksheet("formula-datetime");
		assertEquals("0.1", Ranges.range(sheet, "B40")
				.getFormatText().getCellFormatResult().text);
	}
	
	protected void testSERIESSUM(Book book) {
		Worksheet sheet = book.getWorksheet("formula-math");
		assertEquals("0.7071", Ranges.range(sheet, "B104").getFormatText().getCellFormatResult().text);
	}

	// FIXME
	// #NAME?
	
	protected void testCELL(Book book) {
		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("48", Ranges.range(sheet, "B48").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testINFO(Book book) {
		Worksheet sheet = book.getWorksheet("formula-info");
		assertEquals("12.0", Ranges.range(sheet, "B46").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testASC(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("EXCEL", Ranges.range(sheet, "B2").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testFINDB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B21").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testLEFTB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("\u6771", Ranges.range(sheet, "B28").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testLENB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("11", Ranges.range(sheet, "B32").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testMIDB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("Fluid", Ranges.range(sheet, "B39").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testSEARCHB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("6", Ranges.range(sheet, "B56").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testREPLACEB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("abcde*k", Ranges.range(sheet, "B46").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	protected void testRIGHTB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("Price", Ranges.range(sheet, "B52").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testFORECAST(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("10.61", Ranges.range(sheet, "B67").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testNORMSDIST(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.91", Ranges.range(sheet, "B129").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testCRITBINOM(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B53").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testINTERCEPT(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.05", Ranges.range(sheet, "B94").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testFISHERINV(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.75", Ranges.range(sheet, "B65").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testSTDEVA(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("27.46", Ranges.range(sheet, "B166").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testPERMUT(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("970200", Ranges.range(sheet, "B140").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testLOGINV(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("4.00", Ranges.range(sheet, "B106").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testLINEST(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B103").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testFREQUENCY(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("1", Ranges.range(sheet, "B70").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testGROWTH(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("320196.72", Ranges.range(sheet, "B85").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testFISHER(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.97", Ranges.range(sheet, "B63").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testPERCENTRANK(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.33", Ranges.range(sheet, "B138").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testCORREL(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.9971", Ranges.range(sheet, "B37").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testPERCENTILE(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("1.9", Ranges.range(sheet, "B136").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testSTDEVPA(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("26.05", Ranges.range(sheet, "B170").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testAVERAGEIF(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("14000", Ranges.range(sheet, "B11").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testQUARTILE(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("3.5", Ranges.range(sheet, "B147").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testMORMINV(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("42.00", Ranges.range(sheet, "B127").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testNEGBINOMDIST(Book book) {
		
		Worksheet sheet = book.getWorksheet("formula-statistical");

		assertEquals("0.06", Ranges.range(sheet, "B123").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testCHITEST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.000308", Ranges.range(sheet, "B27").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testBETADIST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.69", Ranges.range(sheet, "B17").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testVARA(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("754.27", Ranges.range(sheet, "B191").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testPROB(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.1", Ranges.range(sheet, "B144").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testZTEST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.09", Ranges.range(sheet, "B199").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testVARPA(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("678.84", Ranges.range(sheet, "B195").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testTTEST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.20", Ranges.range(sheet, "B186").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testTREND(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("133953.33", Ranges.range(sheet, "B179").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testSTEYX(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("3.31", Ranges.range(sheet, "B172").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testFTEST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.65", Ranges.range(sheet, "B73").getFormatText().getCellFormatResult().text);
	}
	
	protected void testFDIST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.01", Ranges.range(sheet, "B59").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testCOVAR(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("5.2", Ranges.range(sheet, "B50").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testLOGNORMDIST(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.04", Ranges.range(sheet, "B108").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testNORMSINV(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("1.33", Ranges.range(sheet, "B131").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testSTANDARDIZE(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("1.33", Ranges.range(sheet, "B162").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testTRIMMEAN(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("3.78", Ranges.range(sheet, "B183").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testCONFIDENCE(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.69", Ranges.range(sheet, "B35").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testRSQ(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.061", Ranges.range(sheet, "B151").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testBETAINV(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("2", Ranges.range(sheet, "B19").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testAVERAGEIFS(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("87.5", Ranges.range(sheet, "B14").getFormatText().getCellFormatResult().text);
	}
	
	// #NAME?
	
	protected void testPEARSON(Book book) {
		Worksheet sheet = book.getWorksheet("formula-statistical");
		assertEquals("0.70", Ranges.range(sheet, "B133").getFormatText().getCellFormatResult().text);
	}

	// #NAME?
	protected void testJIS(Book book) {
		Worksheet sheet = book.getWorksheet("formula-text");
		assertEquals("#NAME?", Ranges.range(sheet, "B75").getFormatText().getCellFormatResult().text);
	}
}

