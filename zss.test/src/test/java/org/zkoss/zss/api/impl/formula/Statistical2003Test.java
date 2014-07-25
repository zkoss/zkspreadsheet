package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Statistical2003Test extends FormulaTestBase {
	
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
	
	// AVEDEV
	@Test
	public void testAVEDEV() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testAVEDEV(book);
	}

	// AVERAGE
	@Test
	public void testAVERAGE() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testAVERAGE(book);
	}
	
	// AVERAGEA
	@Test
	public void testAVERAGEA() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testAVERAGEA(book);
	}

	// BINOMDIST
	@Test
	public void testBINOMDIST() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testBINOMDIST(book);
	}

	// CHIDIST
	@Test
	public void testCHIDIST() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testCHIDIST(book);
	}

	// CHIINV
	@Test
	public void testCHIINV() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testCHIINV(book);
	}

	// COUNT
	@Test
	public void testCOUNT() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testCOUNT(book);
	}

	// COUNTA
	@Test
	public void testCOUNTA() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testCOUNTA(book);
	}

	// COUNTBLANK
	@Test
	public void testCOUNTBLANK() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testCOUNTBLANK(book);
	}

	// COUNTIF
	@Test
	public void testCOUNTIF() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testCOUNTIF(book);
	}

	// DEVSQ
	@Test
	public void testDEVSQ() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testDEVSQ(book);
	}

	// EXPONDIST
	@Test
	public void testEXPONDIST() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testEXPONDIST(book);
	}

	// FDIST
	@Test
	public void testFDIST()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testFDIST(book);
	}
	
	// FINV
	@Test
	public void testFINV()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testFINV(book);
	}
	
	// GAMMADIST
	@Test
	public void testGAMMADIST()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testGAMMADIST(book);
	}
	
	@Test
	public void testGAMMADISTWithCumulative()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testGAMMADISTWithCumulative(book);
	}

	// GAMMAINV1
	@Test
	public void testGAMMAINV() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testGAMMAINV(book);
	}
	
	// GAMMALN
	@Test
	public void testGAMMALN() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testGAMMALN(book);
	}

	// GEOMEAN
	@Test
	public void testGEOMEAN() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testGEOMEAN(book);
	}
	
	// HARMEAN
	@Test
	public void testHARMEAN()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testHARMEAN(book);
	}

	// HYPGEOMDIST
	@Test
	public void testHYPGEOMDIST() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testHYPGEOMDIST(book);
	}

	// KURT
	@Test
	public void testKURT() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testKURT(book);
	}

	// LARGE
	@Test
	public void testLARGE() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testLARGE(book);
	}

	// MAX
	@Test
	public void testMAX() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testMAX(book);
	}
	
	// MAXA
	@Test
	public void testMAXA() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testMAXA(book);
	}

	// MEDIAN
	@Test
	public void testMEDIAN() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testMEDIAN(book);
	}

	// MIN
	@Test
	public void testMIN() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testMIN(book);
	}

	// MINA
	@Test
	public void testMINA() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testMINA(book);
	}

	// MODE
	@Test
	public void testMODE() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testMODE(book);
	}

	// NORMDIST
	@Test
	public void testNORMDIST() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testNORMDIST(book);
	}

	// POISSON
	@Test
	public void testPOISSON() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testPOISSON(book);
	}

	// RANK
	@Test
	public void testRANK() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testRANK(book);
	}

	// SKEW
	@Test
	public void testSKEW() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testSKEW(book);
	}

	// SLOPE
	@Test
	public void testSLOPE() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testSLOPE(book);
	}

	// SMALL
	@Test
	public void testSMALL() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testSMALL(book);
	}

	// STDEV
	@Test
	public void testSTDEV() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testSTDEV(book);
	}
	
	// TDIST
	@Test
	public void testTDISTWithTwoTail()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testTDISTWithTwoTail(book);
	}
	
	@Test
	public void testTDISTWithOneTail()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testTDISTWithOneTail(book);
	}
	
	// TINV
	@Test
	public void testTINV()  {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testTINV(book);
	}

	// VAR
	@Test
	public void testVAR() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testVAR(book);
	}

	// VARP
	@Test
	public void testVARP() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testVARP(book);
	}
	
	// WEIBULL
	@Test
	public void testWEIBULL() {
		Book book = Util.loadBook("TestFile2003-Formula.xls");
		testWEIBULL(book);
	}
}
