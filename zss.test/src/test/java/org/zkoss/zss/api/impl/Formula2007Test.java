package org.zkoss.zss.api.impl;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.model.Book;

/**
 * @author kuro, Hawk
 */
public class Formula2007Test extends FormulaTestBase {

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
	// public void testNOW() {
	// Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
	// Sheet sheet = book.getSheet("formula-datetime");
	// assertEquals("2013/9/11 16:00", Ranges.range(sheet, "B34").getCellData()
	// .getFormatText());

	// }
	// @Test
	// public void testToday() {
	// Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
	// Sheet sheet = book.getSheet("formula-datetime");
	// assertEquals("2013/9/11", Ranges.range(sheet, "B42").getCellData()
	// .getFormatText());
	// }

	@Test
	public void testCOUPNCD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUPNCD(book);
	}

	@Test
	public void testCOUPPCD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUPPCD(book);
	}

	@Test
	public void testWORKDAY() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testWORKDAY(book);
	}

	@Test
	public void testStartDateIsHoliday() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testStartDateIsHoliday(book);
	}

	@Test
	public void testEndDateIsHoliday() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testEndDateIsHoliday(book);
	}

	@Test
	public void testNegativeWorkdayEndDateIsHolday() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNegativeWorkdayEndDateIsHolday(book);
	}

	@Test
	public void testWorkdayBoundary() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testWorkdayBoundary(book);
	}

	@Test
	public void testWorkdaySpecifiedholiday() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testWorkdaySpecifiedholiday(book);
	}

	@Test
	public void testIMREAL() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMREAL(book);
	}

	@Test
	public void testACCRINT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testACCRINT(book);
	}

	@Test
	public void testACCRINTM() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testACCRINTM(book);
	}

	@Test
	public void testAMORDEGRC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testAMORDEGRC(book);
	}

	@Test
	public void testAMORLINC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testAMORLINC(book);
	}

	@Test
	public void testCOUPDAYBS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUPDAYBS(book);
	}

	@Test
	public void testCOUPDAYS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUPDAYS(book);
	}

	@Test
	public void testCOUPDAYSNC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUPDAYSNC(book);
	}

	@Test
	public void testCOUPNUM() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUPNUM(book);
	}

	@Test
	public void testCUMIPMT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCUMIPMT(book);
	}

	@Test
	public void testCUMPRINC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCUMPRINC(book);
	}

	@Test
	public void testDB() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDB(book);
	}

	@Test
	public void testDDB() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDDB(book);
	}

	@Test
	public void testDISC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDISC(book);
	}

	@Test
	public void testDOLLARDE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDOLLARDE(book);
	}

	@Test
	public void testDOLLARFR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDOLLARFR(book);
	}

	@Test
	public void testDURATION() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDURATION(book);
	}

	@Test
	public void testEFFECT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testEFFECT(book);
	}

	@Test
	public void testFV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFV(book);
	}

	@Test
	public void testFVSCHEDULE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFVSCHEDULE(book);
	}

	@Test
	public void testINTRATE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testINTRATE(book);
	}

	@Test
	public void testIPMT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIPMT(book);
	}

	@Test
	public void testIRR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIRR(book);
	}

	@Test
	public void testNOMINAL() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNOMINAL(book);
	}

	@Test
	public void testNPER() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNPER(book);
	}

	@Test
	public void testNPV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNPV(book);
	}

	@Test
	public void testPMT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPMT(book);
	}

	@Test
	public void testPPMT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPPMT(book);
	}

	@Test
	public void testPRICE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPRICE(book);
	}

	@Test
	public void testPRICEDISC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPRICEDISC(book);
	}

	@Test
	public void testPRICEMAT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPRICEMAT(book);
	}

	@Test
	public void testPV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPV(book);
	}

	@Test
	public void testRATE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testRATE(book);
	}

	@Test
	public void testRECEIVED() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testRECEIVED(book);
	}

	@Test
	public void testSLN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSLN(book);
	}

	@Test
	public void testSYD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSYD(book);
	}

	@Test
	public void testTBILLEQ() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTBILLEQ(book);
	}

	@Test
	public void testTBILLPRICE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTBILLPRICE(book);
	}

	@Test
	public void testTBILLYIELD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTBILLYIELD(book);
	}

	@Test
	public void testXNPV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testXNPV(book);
	}

	@Test
	public void testYIELD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testYIELD(book);
	}

	@Test
	public void testYIELDDISC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testYIELDDISC(book);
	}

	@Test
	public void testYIELDMAT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testYIELDMAT(book);
	}

	@Test
	public void testAND() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testAND(book);
	}

	@Test
	public void testFALSE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFALSE(book);
	}

	@Test
	public void testIF() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIF(book);
	}

	@Test
	public void testIFERROR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIFERROR(book);
	}

	@Test
	public void testNOT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNOT(book);
	}

	@Test
	public void testOR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testOR(book);
	}

	@Test
	public void testTRUE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTRUE(book);
	}

	@Test
	public void testABS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testABS(book);
	}

	@Test
	public void testACOS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testACOS(book);
	}

	@Test
	public void testACOSH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testACOSH(book);
	}

	@Test
	public void testASIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testASIN(book);
	}

	@Test
	public void testASINH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testASINH(book);
	}

	@Test
	public void testATAN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testATAN(book);
	}

	@Test
	public void testATAN2() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testATAN2(book);
	}

	@Test
	public void testATANH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testATANH(book);
	}

	@Test
	public void testCEILING() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCEILING(book);
	}

	@Test
	public void testCOMBIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOMBIN(book);
	}

	@Test
	public void testCOS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOS(book);
	}

	@Test
	public void testCOSH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOSH(book);
	}

	@Test
	public void testDEGREES() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDEGREES(book);
	}

	@Test
	public void testEVEN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testEVEN(book);
	}

	@Test
	public void testEXP() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testEXP(book);
	}

	@Test
	public void testFACT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFACT(book);
	}

	@Test
	public void testFACTDOUBLE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFACTDOUBLE(book);
	}

	@Test
	public void testFLOOR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFLOOR(book);
	}

	@Test
	public void testGCD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testGCD(book);
	}

	@Test
	public void testINT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testINT(book);
	}

	@Test
	public void testLCM() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLCM(book);
	}

	@Test
	public void testLN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLN(book);
	}

	@Test
	public void testLOG() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLOG(book);
	}

	@Test
	public void testLOG10() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLOG10(book);
	}

	@Test
	public void testMDETERM() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMDETERM(book);
	}

	@Test
	public void testMINVERSE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMINVERSE(book);
	}

	@Test
	public void testMMULT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMMULT(book);
	}

	@Test
	public void testMOD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMOD(book);
	}

	@Test
	public void testMROUND() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMROUND(book);
	}

	@Test
	public void testMULTINOMIA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMULTINOMIA(book);
	}

	@Test
	public void testODD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testODD(book);
	}

	@Test
	public void testPI() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPI(book);
	}

	@Test
	public void testPOWER() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPOWER(book);
	}

	@Test
	public void testPRODUCT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPRODUCT(book);
	}

	@Test
	public void testQUOTIENT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testQUOTIENT(book);
	}

	@Test
	public void testRADIANS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testRADIANS(book);
	}

	@Test
	public void testROMAN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testROMAN(book);
	}

	@Test
	public void testROUND() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testROUND(book);
	}

	@Test
	public void testROUNDDOWN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testROUNDDOWN(book);
	}

	@Test
	public void testROUNDUP() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testROUNDUP(book);
	}

	@Test
	public void testSIGN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSIGN(book);
	}

	@Test
	public void testSIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSIN(book);
	}

	@Test
	public void testSINH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSINH(book);
	}

	@Test
	public void testSQRT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSQRT(book);
	}

	@Test
	public void testSQRTPI() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSQRTPI(book);
	}

	@Test
	public void testSUBTOTAL() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUBTOTAL(book);
	}

	@Test
	public void testSUM() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUM(book);
	}

	@Test
	public void testSUMIF() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUMIF(book);
	}

	@Test
	public void testSUMIFS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUMIFS(book);
	}

	@Test
	public void testSUMPRODUCT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUMPRODUCT(book);
	}

	@Test
	public void testSUMSQ() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUMSQ(book);
	}

	@Test
	public void testSUMX2MY2() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUMX2MY2(book);
	}

	@Test
	public void testSUMX2PY2() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUMX2PY2(book);
	}

	@Test
	public void testSUMXMY2() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUMXMY2(book);
	}

	@Test
	public void testTAN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTAN(book);
	}

	@Test
	public void testTANH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTANH(book);
	}

	@Test
	public void testTRUNC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTRUNC(book);
	}

	@Test
	public void testDATE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDATE(book);
	}

	@Test
	public void testDATEVALUE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDATEVALUE(book);
	}

	@Test
	public void testDAY() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDAY(book);
	}

	@Test
	public void testDAY360() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDAY360(book);
	}

	@Test
	public void testHOUR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testHOUR(book);
	}

	@Test
	public void testMINUTE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMINUTE(book);
	}

	@Test
	public void testMONTH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMONTH(book);
	}

	@Test
	public void testNETWORKDAYS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNETWORKDAYS(book);
	}

	@Test
	public void testNetworkDaysStartDateIsHoliday() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNetworkDaysStartDateIsHoliday(book);
	}

	@Test
	public void testNetworkDaysAllHoliday() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNetworkDaysAllHoliday(book);
	}

	@Test
	public void testNetworkDaysSpecificHoliday() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNetworkDaysSpecificHoliday(book);
	}

	@Test
	public void testStartDateEqualsEndDate() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testStartDateEqualsEndDate(book);
	}

	@Test
	public void testStartDateLaterThanEndDate() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testStartDateLaterThanEndDate(book);
	}

	@Test
	public void testEmptyDate() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testEmptyDate(book);
	}

	@Test
	public void testSECOND() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSECOND(book);
	}

	@Test
	public void testTIME() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTIME(book);
	}

	@Test
	public void testWEEKDAY() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testWEEKDAY(book);
	}

	@Test
	public void testYEAR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testYEAR(book);
	}

	@Test
	public void testYEARFRAC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testYEARFRAC(book);
	}

	@Test
	public void testCHAR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCHAR(book);
	}

	@Test
	public void testLOWER() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLOWER(book);
	}

	@Test
	public void testCODE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCODE(book);
	}

	@Test
	public void testCLEAN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCLEAN(book);
	}

	@Test
	public void testCONCATENATE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCONCATENATE(book);
	}

	@Test
	public void testEXACT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testEXACT(book);
	}

	@Test
	public void testFIND() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFIND(book);
	}

	@Test
	public void testFIXED() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testFIXED(book);
	}

	@Test
	public void testLEFT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLEFT(book);
	}

	@Test
	public void testLEN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLEN(book);
	}

	@Test
	public void testMID() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMID(book);
	}

	@Test
	public void testPROPER() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPROPER(book);
	}

	@Test
	public void testREPLACE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testREPLACE(book);
	}

	@Test
	public void testREPT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testREPT(book);
	}

	@Test
	public void testRIGHT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testRIGHT(book);
	}

	@Test
	public void testSEARCH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSEARCH(book);
	}

	@Test
	public void testSUBSTITUTE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSUBSTITUTE(book);
	}

	@Test
	public void testT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testT(book);
	}

	@Test
	public void testTEXT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTEXT(book);
	}

	@Test
	public void testTRIM() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTRIM(book);
	}

	@Test
	public void testUPPER() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testUPPER(book);
	}

	@Test
	public void testAVEDEV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testAVEDEV(book);
	}

	@Test
	public void testAVERAGE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testAVERAGE(book);
	}

	@Test
	public void testERRORTYPE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testERRORTYPE(book);
	}

	@Test
	public void testISBLANK() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISBLANK(book);
	}

	@Test
	public void testISERR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISERR(book);
	}

	@Test
	public void testISERROR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISERROR(book);
	}

	@Test
	public void testISEVEN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISEVEN(book);
	}

	@Test
	public void testISLOGICAL() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISLOGICAL(book);
	}

	@Test
	public void testISNA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISNA(book);
	}

	@Test
	public void testISNONTEXT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISNONTEXT(book);
	}

	@Test
	public void testISNUMBER() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISNUMBER(book);
	}

	@Test
	public void testISODD() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISODD(book);
	}

	@Test
	public void testISREF() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISREF(book);
	}

	@Test
	public void testISTEXT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testISTEXT(book);
	}

	@Test
	public void testN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testN(book);
	}

	@Test
	public void testNA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNA(book);
	}

	@Test
	public void testTYPE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testTYPE(book);
	}

	@Test
	public void testAVERAGEA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testAVERAGEA(book);
	}

	@Test
	public void testBIOMDIST() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBIOMDIST(book);
	}

	@Test
	public void testCHIDIST() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCHIDIST(book);
	}

	@Test
	public void testCHIINV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCHIINV(book);
	}

	@Test
	public void testCOUNT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUNT(book);
	}

	@Test
	public void testCOUNTA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUNTA(book);
	}

	@Test
	public void testCOUNTBLANK() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUNTBLANK(book);
	}

	@Test
	public void testCOUNTIF() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOUNTIF(book);
	}

	@Test
	public void testDEVSQ() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDEVSQ(book);
	}

	@Test
	public void testEXPONDIST() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testEXPONDIST(book);
	}

	@Test
	public void testGAMMAINV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testGAMMAINV(book);
	}

	@Test
	public void testGAMMALN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testGAMMALN(book);
	}

	@Test
	public void testGEOMEAN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testGEOMEAN(book);
	}

	@Test
	public void testHYPGEOMDIST() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testHYPGEOMDIST(book);
	}

	@Test
	public void testKURT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testKURT(book);
	}

	@Test
	public void testLARGE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLARGE(book);
	}

	@Test
	public void testMAX() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMAX(book);
	}

	@Test
	public void testMAXA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMAXA(book);
	}

	@Test
	public void testMEDIAN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMEDIAN(book);
	}

	@Test
	public void testMIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMIN(book);
	}

	@Test
	public void testMINA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMINA(book);
	}

	@Test
	public void testMODE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMODE(book);
	}

	@Test
	public void testNORMDIST() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testNORMDIST(book);
	}

	@Test
	public void testPOISSON() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testPOISSON(book);
	}

	@Test
	public void testRANK() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testRANK(book);
	}

	@Test
	public void testSKEW() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSKEW(book);
	}

	@Test
	public void testSLOPE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSLOPE(book);
	}

	@Test
	public void testSMALL() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSMALL(book);
	}

	@Test
	public void testSTDEV() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testSTDEV(book);
	}

	@Test
	public void testVAR() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testVAR(book);
	}

	@Test
	public void testVARP() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testVARP(book);
	}

	@Test
	public void testWEIBULL() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testWEIBULL(book);
	}

	@Test
	public void testBESSELI() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBESSELI(book);
	}

	@Test
	public void testBESSELJ() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBESSELJ(book);
	}

	@Test
	public void testBESSELK() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBESSELK(book);
	}

	@Test
	public void testBESSELY() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBESSELY(book);
	}

	@Test
	public void testBIN2DEC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBIN2DEC(book);
	}

	@Test
	public void testBIN2HEX() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBIN2HEX(book);
	}

	@Test
	public void testBIN2OCT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testBIN2OCT(book);
	}

	@Test
	public void testDEC2BIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDEC2BIN(book);
	}

	@Test
	public void testDEC2HEX() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDEC2HEX(book);
	}

	@Test
	public void testDEC2OCT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDEC2OCT(book);
	}

	@Test
	public void testDELTA() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testDELTA(book);
	}

	@Test
	public void testERF() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testERF(book);
	}

	@Test
	public void testERFC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testERFC(book);
	}

	@Test
	public void testGESTEP() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testGESTEP(book);
	}

	@Test
	public void testHEX2BIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testHEX2BIN(book);
	}

	@Test
	public void testHEX2DEC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testHEX2DEC(book);
	}

	@Test
	public void testHEX2OCT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testHEX2OCT(book);
	}

	@Test
	public void testIMABS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMABS(book);
	}

	@Test
	public void testIMAGINARY() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMAGINARY(book);
	}

	@Test
	public void testIMARGUMENT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMARGUMENT(book);
	}

	@Test
	public void testIMCOS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMCOS(book);
	}

	@Test
	public void testIMLN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMLN(book);

	}

	@Test
	public void testIMSQRT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMSQRT(book);

	}

	@Test
	public void testIMLOG10() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMLOG10(book);

	}

	@Test
	public void testIMSUM() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMSUM(book);

	}

	@Test
	public void testIMSIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMSIN(book);
	}

	@Test
	public void testIMSUB() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMSUB(book);
	}

	@Test
	public void testIMLOG2() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMLOG2(book);
	}

	@Test
	public void testIMPOWER() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testIMPOWER(book);
	}

	@Test
	public void testOCT2BIN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testOCT2BIN(book);
	}

	@Test
	public void testOCT2DEC() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testOCT2DEC(book);
	}

	@Test
	public void testOCT2HEX() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testOCT2HEX(book);
	}

	@Test
	public void testADDRESS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testADDRESS(book);
	}

	@Test
	public void testCHOOSE() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCHOOSE(book);
	}

	@Test
	public void testCOLUMN() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOLUMN(book);
	}

	@Test
	public void testCOLUMNS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testCOLUMNS(book);
	}

	@Test
	public void testHLOOKUP() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testHLOOKUP(book);
	}

	@Test
	public void testHyperLink() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testHyperLink(book);
	}

	@Test
	public void testINDIRECT() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testINDIRECT(book);
	}

	@Test
	public void testLOOKUP() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testLOOKUP(book);
	}

	@Test
	public void testMATCH() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testMATCH(book);
	}

	@Test
	public void testOFFSET() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testOFFSET(book);
	}

	@Test
	public void testROW() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testROW(book);
	}

	@Test
	public void testROWS() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testROWS(book);
	}

	@Test
	public void testVLOOKUP() {
		Book book = Util.loadBook("book/TestFile2007-Formula.xlsx");
		testVLOOKUP(book);
	}

}
