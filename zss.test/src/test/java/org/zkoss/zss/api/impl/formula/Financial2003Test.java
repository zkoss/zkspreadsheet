package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Financial2003Test extends FormulaTestBase {
	
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
	
	// ACCRINT
	@Test
	public void testACCRINT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testACCRINT(book);
	}
	
	// ACCRINTM
	@Test
	public void testACCRINTM() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testACCRINTM(book);
	}
	
	// AMORDEGRC
	@Test
	public void testAMORDEGRC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testAMORDEGRC(book);
	}

	// AMORLINC
	@Test
	public void testAMORLINC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testAMORLINC(book);
	}
	
	// COUPDAYBS
	@Test
	public void testCOUPDAYBS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOUPDAYBS(book);
	}

	// COUPDAYS
	@Test
	public void testCOUPDAYS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOUPDAYS(book);
	}

	// COUPDAYSNC
	@Test
	public void testCOUPDAYSNC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOUPDAYSNC(book);
	}
	
	// COUPNCD
	@Test
	public void testCOUPNCD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOUPNCD(book);
	}
	
	// COUPNUM
	@Test
	public void testCOUPNUM() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOUPNUM(book);
	}

	// COUPPCD
	@Test
	public void testCOUPPCD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOUPPCD(book);
	}

	// CUMIPMT
	@Test
	public void testCUMIPMT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCUMIPMT(book);
	}

	// CUMPRINC
	@Test
	public void testCUMPRINC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCUMPRINC(book);
	}

	// DB
	@Test
	public void testDB() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDB(book);
	}

	// DDB
	@Test
	public void testDDB() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDDB(book);
	}

	// DISC
	@Test
	public void testDISC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDISC(book);
	}

	// DOLLARDE
	@Test
	public void testDOLLARDE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDOLLARDE(book);
	}

	// DOLLARFR
	@Test
	public void testDOLLARFR() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDOLLARFR(book);
	}

	// DURATION
	@Test
	public void testDURATION() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDURATION(book);
	}

	// EFFECT
	@Test
	public void testEFFECT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testEFFECT(book);
	}

	// FV
	@Test
	public void testFV() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testFV(book);
	}

	// FVSCHEDULE
	@Test
	public void testFVSCHEDULE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testFVSCHEDULE(book);
	}

	// INTRATE
	@Test
	public void testINTRATE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testINTRATE(book);
	}

	// IPMT
	@Test
	public void testIPMT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testIPMT(book);
	}

	// IRR
	@Test
	public void testIRR() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testIRR(book);
	}

	// NOMINAL
	@Test
	public void testNOMINAL() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNOMINAL(book);
	}

	// NPER
	@Test
	public void testNPER() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNPER(book);
	}

	// NPV
	@Test
	public void testNPV() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNPV(book);
	}

	// PMT
	@Test
	public void testPMT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPMT(book);
	}

	// PPMT
	@Test
	public void testPPMT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPPMT(book);
	}

	// PRICE
	@Test
	public void testPRICE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPRICE(book);
	}

	// PRICEDISC
	@Test
	public void testPRICEDISC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPRICEDISC(book);
	}

	// PRICEMAT
	@Test
	public void testPRICEMAT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPRICEMAT(book);
	}

	// PV
	@Test
	public void testPV() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPV(book);
	}

	// RATE
	@Test
	public void testRATE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testRATE(book);
	}

	// RECEIVED
	@Test
	public void testRECEIVED() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testRECEIVED(book);
	}

	// SLN
	@Test
	public void testSLN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSLN(book);
	}

	// SYD
	@Test
	public void testSYD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSYD(book);
	}

	// TBILLEQ
	@Test
	public void testTBILLEQ() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTBILLEQ(book);
	}

	// TBILLPRICE
	@Test
	public void testTBILLPRICE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTBILLPRICE(book);
	}

	// TBILLYIELD
	@Test
	public void testTBILLYIELD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTBILLYIELD(book);
	}

	// XNPV
	@Test
	public void testXNPV() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testXNPV(book);
	}

	// YIELD
	@Test
	public void testYIELD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testYIELD(book);
	}

	// YIELDDISC
	@Test
	public void testYIELDDISC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testYIELDDISC(book);
	}

	// YIELDMAT
	@Test
	public void testYIELDMAT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testYIELDMAT(book);
	}
}
