package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Financial2007Test extends FormulaTestBase {
	
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
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testACCRINT(book);
	}
	
	// ACCRINTM
	@Test
	public void testACCRINTM() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testACCRINTM(book);
	}
	
	// AMORDEGRC
	@Test
	public void testAMORDEGRC() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testAMORDEGRC(book);
	}

	// AMORLINC
	@Test
	public void testAMORLINC() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testAMORLINC(book);
	}
	
	// COUPDAYBS
	@Test
	public void testCOUPDAYBS() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCOUPDAYBS(book);
	}

	// COUPDAYS
	@Test
	public void testCOUPDAYS() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCOUPDAYS(book);
	}

	// COUPDAYSNC
	@Test
	public void testCOUPDAYSNC() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCOUPDAYSNC(book);
	}
	
	// COUPNCD
	@Test
	public void testCOUPNCD() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCOUPNCD(book);
	}
	
	// COUPNUM
	@Test
	public void testCOUPNUM() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCOUPNUM(book);
	}

	// COUPPCD
	@Test
	public void testCOUPPCD() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCOUPPCD(book);
	}

	// CUMIPMT
	@Test
	public void testCUMIPMT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCUMIPMT(book);
	}

	// CUMPRINC
	@Test
	public void testCUMPRINC() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testCUMPRINC(book);
	}

	// DB
	@Test
	public void testDB() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testDB(book);
	}

	// DDB
	@Test
	public void testDDB() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testDDB(book);
	}

	// DISC
	@Test
	public void testDISC() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testDISC(book);
	}

	// DOLLARDE
	@Test
	public void testDOLLARDE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testDOLLARDE(book);
	}

	// DOLLARFR
	@Test
	public void testDOLLARFR() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testDOLLARFR(book);
	}

	// DURATION
	@Test
	public void testDURATION() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testDURATION(book);
	}

	// EFFECT
	@Test
	public void testEFFECT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testEFFECT(book);
	}

	// FV
	@Test
	public void testFV() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testFV(book);
	}

	// FVSCHEDULE
	@Test
	public void testFVSCHEDULE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testFVSCHEDULE(book);
	}

	// INTRATE
	@Test
	public void testINTRATE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testINTRATE(book);
	}

	// IPMT
	@Test
	public void testIPMT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testIPMT(book);
	}

	// IRR
	@Test
	public void testIRR() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testIRR(book);
	}

	// NOMINAL
	@Test
	public void testNOMINAL() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testNOMINAL(book);
	}

	// NPER
	@Test
	public void testNPER() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testNPER(book);
	}

	// NPV
	@Test
	public void testNPV() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testNPV(book);
	}

	// PMT
	@Test
	public void testPMT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testPMT(book);
	}

	// PPMT
	@Test
	public void testPPMT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testPPMT(book);
	}

	// PRICE
	@Test
	public void testPRICE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testPRICE(book);
	}

	// PRICEDISC
	@Test
	public void testPRICEDISC() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testPRICEDISC(book);
	}

	// PRICEMAT
	@Test
	public void testPRICEMAT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testPRICEMAT(book);
	}

	// PV
	@Test
	public void testPV() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testPV(book);
	}

	// RATE
	@Test
	public void testRATE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testRATE(book);
	}

	// RECEIVED
	@Test
	public void testRECEIVED() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testRECEIVED(book);
	}

	// SLN
	@Test
	public void testSLN() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testSLN(book);
	}

	// SYD
	@Test
	public void testSYD() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testSYD(book);
	}

	// TBILLEQ
	@Test
	public void testTBILLEQ() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testTBILLEQ(book);
	}

	// TBILLPRICE
	@Test
	public void testTBILLPRICE() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testTBILLPRICE(book);
	}

	// TBILLYIELD
	@Test
	public void testTBILLYIELD() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testTBILLYIELD(book);
	}

	// XNPV
	@Test
	public void testXNPV() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testXNPV(book);
	}

	// YIELD
	@Test
	public void testYIELD() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testYIELD(book);
	}

	// YIELDDISC
	@Test
	public void testYIELDDISC() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testYIELDDISC(book);
	}

	// YIELDMAT
	@Test
	public void testYIELDMAT() {
		Book book = Util.loadBook(this, "book/TestFile2007-Formula.xlsx");
		testYIELDMAT(book);
	}
}
