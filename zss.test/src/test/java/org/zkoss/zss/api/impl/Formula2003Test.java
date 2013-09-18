package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author kuro, Hawk
 */
public class Formula2003Test extends FormulaTestBase {

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

	@Test
	public void testACCRINT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testACCRINT(book);
	}

	@Test
	public void testACCRINTM() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testACCRINTM(book);
	}

	@Test
	public void testAMORDEGRC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testAMORDEGRC(book);
	}

	@Test
	public void testAMORLINC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testAMORLINC(book);
	}

	@Test
	public void testCOUPDAYBS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUPDAYBS(book);
	}

	@Test
	public void testCOUPDAYS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUPDAYS(book);
	}

	@Test
	public void testCOUPDAYSNC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUPDAYSNC(book);
	}

	@Test
	public void testCOUPNUM() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUPNUM(book);
	}

	@Test
	public void testCUMIPMT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCUMIPMT(book);
	}

	@Test
	public void testCUMPRINC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCUMPRINC(book);
	}

	@Test
	public void testDB() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDB(book);
	}

	@Test
	public void testDDB() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDDB(book);
	}

	@Test
	public void testDISC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDISC(book);
	}

	@Test
	public void testDOLLARDE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDOLLARDE(book);
	}

	@Test
	public void testDOLLARFR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDOLLARFR(book);
	}

	@Test
	public void testDURATION() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDURATION(book);
	}

	@Test
	public void testEFFECT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testEFFECT(book);
	}

	@Test
	public void testFV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFV(book);
	}

	@Test
	public void testFVSCHEDULE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFVSCHEDULE(book);
	}

	@Test
	public void testINTRATE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testINTRATE(book);
	}

	@Test
	public void testIPMT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIPMT(book);
	}

	@Test
	public void testIRR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIRR(book);
	}

	@Test
	public void testNOMINAL() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNOMINAL(book);
	}

	@Test
	public void testNPER() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNPER(book);
	}

	@Test
	public void testNPV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNPV(book);
	}

	@Test
	public void testPMT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPMT(book);
	}

	@Test
	public void testPPMT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPPMT(book);
	}

	@Test
	public void testPRICE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPRICE(book);
	}

	@Test
	public void testPRICEDISC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPRICEDISC(book);
	}

	@Test
	public void testPRICEMAT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPRICEMAT(book);
	}

	@Test
	public void testPV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPV(book);
	}

	@Test
	public void testRATE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testRATE(book);
	}

	@Test
	public void testRECEIVED() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testRECEIVED(book);
	}

	@Test
	public void testSLN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSLN(book);
	}

	@Test
	public void testSYD() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSYD(book);
	}

	@Test
	public void testTBILLEQ() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTBILLEQ(book);
	}

	@Test
	public void testTBILLPRICE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTBILLPRICE(book);
	}

	@Test
	public void testTBILLYIELD() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTBILLYIELD(book);
	}

	@Test
	public void testXNPV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testXNPV(book);
	}

	@Test
	public void testYIELD() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testYIELD(book);
	}

	@Test
	public void testYIELDDISC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testYIELDDISC(book);
	}

	@Test
	public void testYIELDMAT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testYIELDMAT(book);
	}

	@Test
	public void testAND() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testAND(book);
	}

	@Test
	public void testFALSE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFALSE(book);
	}

	@Test
	public void testIF() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIF(book);
	}

	@Test
	public void testIFERROR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIFERROR(book);
	}

	@Test
	public void testNOT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNOT(book);
	}

	@Test
	public void testOR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testOR(book);
	}

	@Test
	public void testTRUE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTRUE(book);
	}

	@Test
	public void testABS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testABS(book);
	}

	@Test
	public void testACOS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testACOS(book);
	}

	@Test
	public void testACOSH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testACOSH(book);
	}

	@Test
	public void testASIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testASIN(book);
	}

	@Test
	public void testASINH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testASINH(book);
	}

	@Test
	public void testATAN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testATAN(book);
	}

	@Test
	public void testATAN2() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testATAN2(book);
	}

	@Test
	public void testATANH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testATANH(book);
	}

	@Test
	public void testCEILING() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCEILING(book);
	}

	@Test
	public void testCOMBIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOMBIN(book);
	}

	@Test
	public void testCOS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOS(book);
	}

	@Test
	public void testCOSH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOSH(book);
	}

	@Test
	public void testDEGREES() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDEGREES(book);
	}

	@Test
	public void testEVEN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testEVEN(book);
	}

	@Test
	public void testEXP() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testEXP(book);
	}

	@Test
	public void testFACT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFACT(book);
	}

	@Test
	public void testFACTDOUBLE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFACTDOUBLE(book);
	}

	@Test
	public void testFLOOR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFLOOR(book);
	}

	@Test
	public void testGCD() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testGCD(book);
	}

	@Test
	public void testINT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testINT(book);
	}

	@Test
	public void testLCM() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLCM(book);
	}

	@Test
	public void testLN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLN(book);
	}

	@Test
	public void testLOG() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLOG(book);
	}

	@Test
	public void testLOG10() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLOG10(book);
	}

	@Test
	public void testMDETERM() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMDETERM(book);
	}

	@Test
	public void testMINVERSE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMINVERSE(book);
	}

	@Test
	public void testMMULT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMMULT(book);
	}

	@Test
	public void testMOD() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMOD(book);
	}

	@Test
	public void testMROUND() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMROUND(book);
	}

	@Test
	public void testMULTINOMIA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMULTINOMIA(book);
	}

	@Test
	public void testODD() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testODD(book);
	}

	@Test
	public void testPI() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPI(book);
	}

	@Test
	public void testPOWER() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPOWER(book);
	}

	@Test
	public void testPRODUCT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPRODUCT(book);
	}

	@Test
	public void testQUOTIENT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testQUOTIENT(book);
	}

	@Test
	public void testRADIANS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testRADIANS(book);
	}

	@Test
	public void testROMAN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testROMAN(book);
	}

	@Test
	public void testROUND() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testROUND(book);
	}

	@Test
	public void testROUNDDOWN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testROUNDDOWN(book);
	}

	@Test
	public void testROUNDUP() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testROUNDUP(book);
	}

	@Test
	public void testSIGN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSIGN(book);
	}

	@Test
	public void testSIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSIN(book);
	}

	@Test
	public void testSINH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSINH(book);
	}

	@Test
	public void testSQRT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSQRT(book);
	}

	@Test
	public void testSQRTPI() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSQRTPI(book);
	}

	@Test
	public void testSUBTOTAL() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUBTOTAL(book);
	}

	@Test
	public void testSUM() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUM(book);
	}

	@Test
	public void testSUMIF() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUMIF(book);
	}

	@Test
	public void testSUMIFS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUMIFS(book);
	}

	@Test
	public void testSUMPRODUCT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUMPRODUCT(book);
	}

	@Test
	public void testSUMSQ() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUMSQ(book);
	}

	@Test
	public void testSUMX2MY2() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUMX2MY2(book);
	}

	@Test
	public void testSUMX2PY2() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUMX2PY2(book);
	}

	@Test
	public void testSUMXMY2() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUMXMY2(book);
	}

	@Test
	public void testTAN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTAN(book);
	}

	@Test
	public void testTANH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		Sheet sheet = book.getSheet("formula-math");
		assertEquals("-0.96", Ranges.range(sheet, "B182").getCellData().getFormatText());
	}

	@Test
	public void testTRUNC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTRUNC(book);
	}

	@Test
	public void testDATE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDATE(book);
	}

	@Test
	public void testDATEVALUE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDATEVALUE(book);
	}

	@Test
	public void testDAY() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDAY(book);
	}

	@Test
	public void testDAY360() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDAY360(book);
	}

	@Test
	public void testHOUR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testHOUR(book);
	}

	@Test
	public void testMINUTE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMINUTE(book);
	}

	@Test
	public void testMONTH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMONTH(book);
	}

	@Test
	public void testNETWORKDAYS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNETWORKDAYS(book);
	}

	@Test
	public void testNetworkDaysStartDateIsHoliday() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNetworkDaysStartDateIsHoliday(book);
	}

	@Test
	public void testNetworkDaysAllHoliday() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNetworkDaysAllHoliday(book);
	}

	@Test
	public void testNetworkDaysSpecificHoliday() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNetworkDaysSpecificHoliday(book);
	}

	@Test
	public void testStartDateEqualsEndDate() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testStartDateEqualsEndDate(book);
	}

	@Test
	public void testStartDateLaterThanEndDate() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testStartDateLaterThanEndDate(book);
	}

	@Test
	public void testEmptyDate() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testEmptyDate(book);
	}

	@Test
	public void testSECOND() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSECOND(book);
	}

	@Test
	public void testTIME() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTIME(book);
	}

	@Test
	public void testWEEKDAY() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testWEEKDAY(book);
	}

	@Test
	public void testYEAR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testYEAR(book);
	}

	@Test
	public void testYEARFRAC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testYEARFRAC(book);
	}

	@Test
	public void testCHAR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCHAR(book);
	}

	@Test
	public void testLOWER() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLOWER(book);
	}

	@Test
	public void testCODE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCODE(book);
	}

	@Test
	public void testCLEAN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCLEAN(book);
	}

	@Test
	public void testCONCATENATE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCONCATENATE(book);
	}

	@Test
	public void testEXACT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testEXACT(book);
	}

	@Test
	public void testFIND() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFIND(book);
	}

	@Test
	public void testFIXED() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testFIXED(book);
	}

	@Test
	public void testLEFT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLEFT(book);
	}

	@Test
	public void testLEN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLEN(book);
	}

	@Test
	public void testMID() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMID(book);
	}

	@Test
	public void testPROPER() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPROPER(book);
	}

	@Test
	public void testREPLACE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testREPLACE(book);
	}

	@Test
	public void testREPT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testREPT(book);
	}

	@Test
	public void testRIGHT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testRIGHT(book);
	}

	@Test
	public void testSEARCH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSEARCH(book);
	}

	@Test
	public void testSUBSTITUTE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSUBSTITUTE(book);
	}

	@Test
	public void testT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testT(book);
	}

	@Test
	public void testTEXT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTEXT(book);
	}

	@Test
	public void testTRIM() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTRIM(book);
	}

	@Test
	public void testUPPER() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testUPPER(book);
	}

	@Test
	public void testAVEDEV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testAVEDEV(book);
	}

	@Test
	public void testAVERAGE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testAVERAGE(book);
	}

	@Test
	public void testERRORTYPE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testERRORTYPE(book);
	}

	@Test
	public void testISBLANK() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISBLANK(book);
	}

	@Test
	public void testISERR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISERR(book);
	}

	@Test
	public void testISERROR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISERROR(book);
	}

	@Test
	public void testISEVEN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISEVEN(book);
	}

	@Test
	public void testISLOGICAL() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISLOGICAL(book);
	}

	@Test
	public void testISNA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISNA(book);
	}

	@Test
	public void testISNONTEXT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISNONTEXT(book);
	}

	@Test
	public void testISNUMBER() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISNUMBER(book);
	}

	@Test
	public void testISODD() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISODD(book);
	}

	@Test
	public void testISREF() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISREF(book);
	}

	@Test
	public void testISTEXT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testISTEXT(book);
	}

	@Test
	public void testN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testN(book);
	}

	@Test
	public void testNA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNA(book);
	}

	@Test
	public void testTYPE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testTYPE(book);
	}

	@Test
	public void testAVERAGEA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testAVERAGEA(book);
	}

	@Test
	public void testBIOMDIST() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBIOMDIST(book);
	}

	@Test
	public void testCHIDIST() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCHIDIST(book);
	}

	@Test
	public void testCHIINV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCHIINV(book);
	}

	@Test
	public void testCOUNT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUNT(book);
	}

	@Test
	public void testCOUNTA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUNTA(book);
	}

	@Test
	public void testCOUNTBLANK() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUNTBLANK(book);
	}

	@Test
	public void testCOUNTIF() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOUNTIF(book);
	}

	@Test
	public void testDEVSQ() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDEVSQ(book);
	}

	@Test
	public void testEXPONDIST() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testEXPONDIST(book);
	}

	@Test
	public void testGAMMAINV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testGAMMAINV(book);
	}

	@Test
	public void testGAMMALN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testGAMMALN(book);
	}

	@Test
	public void testGEOMEAN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testGEOMEAN(book);
	}

	@Test
	public void testHYPGEOMDIST() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testHYPGEOMDIST(book);
	}

	@Test
	public void testKURT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testKURT(book);
	}

	@Test
	public void testLARGE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLARGE(book);
	}

	@Test
	public void testMAX() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMAX(book);
	}

	@Test
	public void testMAXA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMAXA(book);
	}

	@Test
	public void testMEDIAN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMEDIAN(book);
	}

	@Test
	public void testMIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMIN(book);
	}

	@Test
	public void testMINA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMINA(book);
	}

	@Test
	public void testMODE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMODE(book);
	}

	@Test
	public void testNORMDIST() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testNORMDIST(book);
	}

	@Test
	public void testPOISSON() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testPOISSON(book);
	}

	@Test
	public void testRANK() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testRANK(book);
	}

	@Test
	public void testSKEW() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSKEW(book);
	}

	@Test
	public void testSLOPE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSLOPE(book);
	}

	@Test
	public void testSMALL() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSMALL(book);
	}

	@Test
	public void testSTDEV() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testSTDEV(book);
	}

	@Test
	public void testVAR() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testVAR(book);
	}

	@Test
	public void testVARP() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testVARP(book);
	}

	@Test
	public void testWEIBULL() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testWEIBULL(book);
	}

	@Test
	public void testBESSELI() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBESSELI(book);
	}

	@Test
	public void testBESSELJ() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBESSELJ(book);
	}

	@Test
	public void testBESSELK() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBESSELK(book);
	}

	@Test
	public void testBESSELY() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBESSELY(book);
	}

	@Test
	public void testBIN2DEC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBIN2DEC(book);
	}

	@Test
	public void testBIN2HEX() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBIN2HEX(book);
	}

	@Test
	public void testBIN2OCT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testBIN2OCT(book);
	}

	@Test
	public void testDEC2BIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDEC2BIN(book);
	}

	@Test
	public void testDEC2HEX() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDEC2HEX(book);
	}

	@Test
	public void testDEC2OCT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDEC2OCT(book);
	}

	@Test
	public void testDELTA() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testDELTA(book);
	}

	@Test
	public void testERF() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testERF(book);
	}

	@Test
	public void testERFC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testERFC(book);
	}

	@Test
	public void testGESTEP() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testGESTEP(book);
	}

	@Test
	public void testHEX2BIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testHEX2BIN(book);
	}

	@Test
	public void testHEX2DEC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testHEX2DEC(book);
	}

	@Test
	public void testHEX2OCT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testHEX2OCT(book);
	}

	@Test
	public void testIMABS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMABS(book);
	}

	@Test
	public void testIMAGINARY() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMAGINARY(book);
	}

	@Test
	public void testIMARGUMENT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMARGUMENT(book);
	}

	@Test
	public void testIMCOS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMCOS(book);
	}

	@Test
	public void testIMLN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMLN(book);

	}

	@Test
	public void testIMSQRT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMSQRT(book);

	}

	@Test
	public void testIMLOG10() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMLOG10(book);

	}

	@Test
	public void testIMSUM() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMSUM(book);

	}

	@Test
	public void testIMSIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMSIN(book);
	}

	@Test
	public void testIMSUB() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMSUB(book);
	}

	@Test
	public void testIMLOG2() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMLOG2(book);
	}

	@Test
	public void testIMPOWER() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testIMPOWER(book);
	}
	
	@Test
	public void testOCT2BIN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testOCT2BIN(book);
	}

	@Test
	public void testOCT2DEC() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testOCT2DEC(book);
	}

	@Test
	public void testOCT2HEX() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testOCT2HEX(book);
	}

	@Test
	public void testADDRESS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testADDRESS(book);
	}

	@Test
	public void testCHOOSE() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCHOOSE(book);
	}

	@Test
	public void testCOLUMN() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOLUMN(book);
	}

	@Test
	public void testCOLUMNS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testCOLUMNS(book);
	}

	@Test
	public void testHLOOKUP() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testHLOOKUP(book);
	}

	@Test
	public void testHyperLink() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testHyperLink(book);
	}

	@Test
	public void testINDIRECT() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testINDIRECT(book);
	}

	@Test
	public void testLOOKUP() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testLOOKUP(book);
	}

	@Test
	public void testMATCH() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testMATCH(book);
	}

	@Test
	public void testOFFSET() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testOFFSET(book);
	}

	@Test
	public void testROW() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testROW(book);
	}

	@Test
	public void testROWS() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testROWS(book);
	}

	@Test
	public void testVLOOKUP() {
		Book book = Util.loadBook("book/TestFile2003-Formula.xls");
		testVLOOKUP(book);
	}

}
