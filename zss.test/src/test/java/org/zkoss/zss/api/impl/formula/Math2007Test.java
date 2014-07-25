package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Math2007Test extends FormulaTestBase {
	
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
	
	// ABS
	@Test
	public void testABS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testABS(book);
	}

	// ACOS
	@Test
	public void testACOS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testACOS(book);
	}

	// ACOSH
	@Test
	public void testACOSH() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testACOSH(book);
	}

	// ASIN
	@Test
	public void testASIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testASIN(book);
	}

	// ASINH
	@Test
	public void testASINH() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testASINH(book);
	}

	// ATAN
	@Test
	public void testATAN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testATAN(book);
	}

	// ATAN2
	@Test
	public void testATAN2() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testATAN2(book);
	}

	// ATANH
	@Test
	public void testATANH() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testATANH(book);
	}

	// CEILING
	@Test
	public void testCEILING() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCEILING(book);
	}

	// COMBIN
	@Test
	public void testCOMBIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOMBIN(book);
	}

	// COS
	@Test
	public void testCOS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOS(book);
	}

	// COSH
	@Test
	public void testCOSH() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOSH(book);
	}

	// DEGREES
	@Test
	public void testDEGREES() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testDEGREES(book);
	}

	// EVEN
	@Test
	public void testEVEN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testEVEN(book);
	}

	// EXP
	@Test
	public void testEXP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testEXP(book);
	}

	// FACT
	@Test
	public void testFACT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testFACT(book);
	}

	// FACTDOUBLE
	@Test
	public void testFACTDOUBLE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testFACTDOUBLE(book);
	}

	// FLOOR
	@Test
	public void testFLOOR() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testFLOOR(book);
	}

	// GCD
	@Test
	public void testGCD() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testGCD(book);
	}

	// INT
	@Test
	public void testINT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testINT(book);
	}

	// LCM
	@Test
	public void testLCM() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testLCM(book);
	}

	// LN
	@Test
	public void testLN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testLN(book);
	}

	// LOG
	@Test
	public void testLOG() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testLOG(book);
	}

	// LOG10
	@Test
	public void testLOG10() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testLOG10(book);
	}

	// MIDETERM
	@Test
	public void testMDETERM() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMDETERM(book);
	}

	// MINVERSE
	@Test
	public void testMINVERSE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMINVERSE(book);
	}

	// MMULT
	@Test
	public void testMMULT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMMULT(book);
	}

	// MDD
	@Test
	public void testMOD() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMOD(book);
	}

	// MROUND
	@Test
	public void testMROUND() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMROUND(book);
	}

	// MULTINOMIA
	@Test
	public void testMULTINOMIA() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMULTINOMIA(book);
	}

	// ODD
	@Test
	public void testODD() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testODD(book);
	}

	// PI
	@Test
	public void testPI() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testPI(book);
	}

	// POWER
	@Test
	public void testPOWER() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testPOWER(book);
	}

	// PRODUCT
	@Test
	public void testPRODUCT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testPRODUCT(book);
	}

	// QUOTIENT
	@Test
	public void testQUOTIENT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testQUOTIENT(book);
	}

	// RADIANS
	@Test
	public void testRADIANS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testRADIANS(book);
	}

	// ROMAN
	@Test
	public void testROMAN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testROMAN(book);
	}

	// ROUND
	@Test
	public void testROUND() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testROUND(book);
	}

	// ROUNDDOWN
	@Test
	public void testROUNDDOWN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testROUNDDOWN(book);
	}

	// ROUNDUP
	@Test
	public void testROUNDUP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testROUNDUP(book);
	}

	// SIGN
	@Test
	public void testSIGN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSIGN(book);
	}

	// SIN
	@Test
	public void testSIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSIN(book);
	}

	// SINH
	@Test
	public void testSINH() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSINH(book);
	}

	// SQRT
	@Test
	public void testSQRT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSQRT(book);
	}

	// SQRTPI
	@Test
	public void testSQRTPI() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSQRTPI(book);
	}

	// SUBTOTAL
	@Test
	public void testSUBTOTAL() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUBTOTAL(book);
	}

	// SUM
	@Test
	public void testSUM() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUM(book);
	}
	
	// SUMIF
	@Test
	public void testSUMIF() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUMIF(book);
	}

	// SUMIFS
	@Test
	public void testSUMIFS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUMIFS(book);
	}

	// SUMPRODUCT
	@Test
	public void testSUMPRODUCT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUMPRODUCT(book);
	}

	// SUMSQ
	@Test
	public void testSUMSQ() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUMSQ(book);
	}

	// SUMX2MY2
	@Test
	public void testSUMX2MY2() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUMX2MY2(book);
	}

	// SUMX2PY2
	@Test
	public void testSUMX2PY2() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUMX2PY2(book);
	}

	// SUMXMY2
	@Test
	public void testSUMXMY2() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSUMXMY2(book);
	}

	// TAN
	@Test
	public void testTAN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testTAN(book);
	}

	// TANH
	@Test
	public void testTANH() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testTANH(book);
	}

	// TRUNC
	@Test
	public void testTRUNC() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testTRUNC(book);
	}
}
