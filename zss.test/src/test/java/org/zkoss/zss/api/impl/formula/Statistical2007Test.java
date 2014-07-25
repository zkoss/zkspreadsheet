package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Statistical2007Test extends FormulaTestBase {
	
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
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testAVEDEV(book);
	}

	// AVERAGE
	@Test
	public void testAVERAGE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testAVERAGE(book);
	}
	
	// AVERAGEA
	@Test
	public void testAVERAGEA() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testAVERAGEA(book);
	}

	// BINOMDIST
	@Test
	public void testBINOMDIST() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBINOMDIST(book);
	}

	// CHIDIST
	@Test
	public void testCHIDIST() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCHIDIST(book);
	}

	// CHIINV
	@Test
	public void testCHIINV() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCHIINV(book);
	}

	// COUNT
	@Test
	public void testCOUNT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOUNT(book);
	}

	// COUNTA
	@Test
	public void testCOUNTA() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOUNTA(book);
	}

	// COUNTBLANK
	@Test
	public void testCOUNTBLANK() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOUNTBLANK(book);
	}

	// COUNTIF
	@Test
	public void testCOUNTIF() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOUNTIF(book);
	}

	// DEVSQ
	@Test
	public void testDEVSQ() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testDEVSQ(book);
	}

	// EXPONDIST
	@Test
	public void testEXPONDIST() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testEXPONDIST(book);
	}

	// FDIST
	@Test
	public void testFDIST()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testFDIST(book);
	}
	
	// FINV
	@Test
	public void testFINV()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testFINV(book);
	}
	
	// GAMMADIST
	@Test
	public void testGAMMADIST()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testGAMMADIST(book);
	}
	
	@Test
	public void testGAMMADISTWithCumulative()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testGAMMADISTWithCumulative(book);
	}

	// GAMMAINV1
	@Test
	public void testGAMMAINV() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testGAMMAINV(book);
	}
	
	// GAMMALN
	@Test
	public void testGAMMALN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testGAMMALN(book);
	}

	// GEOMEAN
	@Test
	public void testGEOMEAN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testGEOMEAN(book);
	}
	
	// HARMEAN
	@Test
	public void testHARMEAN()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testHARMEAN(book);
	}

	// HYPGEOMDIST
	@Test
	public void testHYPGEOMDIST() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testHYPGEOMDIST(book);
	}

	// KURT
	@Test
	public void testKURT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testKURT(book);
	}

	// LARGE
	@Test
	public void testLARGE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testLARGE(book);
	}

	// MAX
	@Test
	public void testMAX() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMAX(book);
	}
	
	// MAXA
	@Test
	public void testMAXA() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMAXA(book);
	}

	// MEDIAN
	@Test
	public void testMEDIAN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMEDIAN(book);
	}

	// MIN
	@Test
	public void testMIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMIN(book);
	}

	// MINA
	@Test
	public void testMINA() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMINA(book);
	}

	// MODE
	@Test
	public void testMODE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testMODE(book);
	}

	// NORMDIST
	@Test
	public void testNORMDIST() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testNORMDIST(book);
	}

	// POISSON
	@Test
	public void testPOISSON() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testPOISSON(book);
	}

	// RANK
	@Test
	public void testRANK() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testRANK(book);
	}

	// SKEW
	@Test
	public void testSKEW() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSKEW(book);
	}

	// SLOPE
	@Test
	public void testSLOPE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSLOPE(book);
	}

	// SMALL
	@Test
	public void testSMALL() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSMALL(book);
	}

	// STDEV
	@Test
	public void testSTDEV() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testSTDEV(book);
	}
	
	// TDIST
	@Test
	public void testTDISTWithTwoTail()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testTDISTWithTwoTail(book);
	}
	
	@Test
	public void testTDISTWithOneTail()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testTDISTWithOneTail(book);
	}
	
	// TINV
	@Test
	public void testTINV()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testTINV(book);
	}

	// VAR
	@Test
	public void testVAR() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testVAR(book);
	}

	// VARP
	@Test
	public void testVARP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testVARP(book);
	}
	
	// WEIBULL
	@Test
	public void testWEIBULL() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testWEIBULL(book);
	}
}
