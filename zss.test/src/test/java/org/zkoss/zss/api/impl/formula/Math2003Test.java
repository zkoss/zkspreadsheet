package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Math2003Test extends FormulaTestBase {
	
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
	
	// ABS
	@Test
	public void testABS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testABS(book);
	}

	// ACOS
	@Test
	public void testACOS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testACOS(book);
	}

	// ACOSH
	@Test
	public void testACOSH() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testACOSH(book);
	}

	// ASIN
	@Test
	public void testASIN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testASIN(book);
	}

	// ASINH
	@Test
	public void testASINH() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testASINH(book);
	}

	// ATAN
	@Test
	public void testATAN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testATAN(book);
	}

	// ATAN2
	@Test
	public void testATAN2() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testATAN2(book);
	}

	// ATANH
	@Test
	public void testATANH() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testATANH(book);
	}

	// CEILING
	@Test
	public void testCEILING() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCEILING(book);
	}

	// COMBIN
	@Test
	public void testCOMBIN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOMBIN(book);
	}

	// COS
	@Test
	public void testCOS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOS(book);
	}

	// COSH
	@Test
	public void testCOSH() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testCOSH(book);
	}

	// DEGREES
	@Test
	public void testDEGREES() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDEGREES(book);
	}

	// EVEN
	@Test
	public void testEVEN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testEVEN(book);
	}

	// EXP
	@Test
	public void testEXP() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testEXP(book);
	}

	// FACT
	@Test
	public void testFACT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testFACT(book);
	}

	// FACTDOUBLE
	@Test
	public void testFACTDOUBLE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testFACTDOUBLE(book);
	}

	// FLOOR
	@Test
	public void testFLOOR() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testFLOOR(book);
	}

	// GCD
	@Test
	public void testGCD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testGCD(book);
	}

	// INT
	@Test
	public void testINT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testINT(book);
	}

	// LCM
	@Test
	public void testLCM() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testLCM(book);
	}

	// LN
	@Test
	public void testLN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testLN(book);
	}

	// LOG
	@Test
	public void testLOG() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testLOG(book);
	}

	// LOG10
	@Test
	public void testLOG10() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testLOG10(book);
	}

	// MIDETERM
	@Test
	public void testMDETERM() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMDETERM(book);
	}

	// MINVERSE
	@Test
	public void testMINVERSE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMINVERSE(book);
	}

	// MMULT
	@Test
	public void testMMULT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMMULT(book);
	}

	// MDD
	@Test
	public void testMOD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMOD(book);
	}

	// MROUND
	@Test
	public void testMROUND() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMROUND(book);
	}

	// MULTINOMIA
	@Test
	public void testMULTINOMIA() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMULTINOMIA(book);
	}

	// ODD
	@Test
	public void testODD() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testODD(book);
	}

	// PI
	@Test
	public void testPI() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPI(book);
	}

	// POWER
	@Test
	public void testPOWER() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPOWER(book);
	}

	// PRODUCT
	@Test
	public void testPRODUCT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testPRODUCT(book);
	}

	// QUOTIENT
	@Test
	public void testQUOTIENT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testQUOTIENT(book);
	}

	// RADIANS
	@Test
	public void testRADIANS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testRADIANS(book);
	}

	// ROMAN
	@Test
	public void testROMAN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testROMAN(book);
	}

	// ROUND
	@Test
	public void testROUND() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testROUND(book);
	}

	// ROUNDDOWN
	@Test
	public void testROUNDDOWN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testROUNDDOWN(book);
	}

	// ROUNDUP
	@Test
	public void testROUNDUP() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testROUNDUP(book);
	}

	// SIGN
	@Test
	public void testSIGN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSIGN(book);
	}

	// SIN
	@Test
	public void testSIN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSIN(book);
	}

	// SINH
	@Test
	public void testSINH() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSINH(book);
	}

	// SQRT
	@Test
	public void testSQRT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSQRT(book);
	}

	// SQRTPI
	@Test
	public void testSQRTPI() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSQRTPI(book);
	}

	// SUBTOTAL
	@Test
	public void testSUBTOTAL() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUBTOTAL(book);
	}

	// SUM
	@Test
	public void testSUM() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUM(book);
	}
	
	// SUMIF
	@Test
	public void testSUMIF() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUMIF(book);
	}

	// SUMIFS
	@Test
	public void testSUMIFS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUMIFS(book);
	}

	// SUMPRODUCT
	@Test
	public void testSUMPRODUCT() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUMPRODUCT(book);
	}

	// SUMSQ
	@Test
	public void testSUMSQ() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUMSQ(book);
	}

	// SUMX2MY2
	@Test
	public void testSUMX2MY2() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUMX2MY2(book);
	}

	// SUMX2PY2
	@Test
	public void testSUMX2PY2() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUMX2PY2(book);
	}

	// SUMXMY2
	@Test
	public void testSUMXMY2() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSUMXMY2(book);
	}

	// TAN
	@Test
	public void testTAN() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTAN(book);
	}

	// TANH
	@Test
	public void testTANH() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTANH(book);
	}

	// TRUNC
	@Test
	public void testTRUNC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTRUNC(book);
	}
}
