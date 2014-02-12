package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

@Ignore
public class StatisticalUnsupportedTest extends FormulaTestBase {
	
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
	
	// #NAME?
	@Test
	public void testSTDEVP()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testSTDEVP(book);
	}
	
	// #NAME?
	@Test
	public void testFORECAST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testFORECAST(book);
	}
	
	// #NAME?
	@Test
	public void testNORMSDIST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testNORMSDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCRITBINOM()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testCRITBINOM(book);
	}
	
	// #NAME?
	@Test
	public void testINTERCEPT()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testINTERCEPT(book);
	}
	
	// #NAME?
	@Test
	public void testFISHERINV()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testFISHERINV(book);
	}
	
	// #NAME?
	@Test
	public void testSTDEVA()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testSTDEVA(book);
	}
	
	// #NAME?
	@Test
	public void testPERMUT()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testPERMUT(book);
	}
	
	// #NAME?
	@Test
	public void testLOGINV()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testLOGINV(book);
	}
	
	// #NAME?
	@Test
	public void testLINEST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testLINEST(book);
	}
	
	// #NAME?
	@Test
	public void testFREQUENCY()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testFREQUENCY(book);
	}
	
	// #NAME?
	@Test
	public void testGROWTH()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testGROWTH(book);
	}
	
	// #NAME?
	@Test
	public void testFISHER()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testFISHER(book);
	}
	
	// #NAME?
	@Test
	public void testPERCENTRANK()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testPERCENTRANK(book);
	}
	
	// #NAME?
	@Test
	public void testCORREL()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testCORREL(book);
	}
	
	// #NAME?
	@Test
	public void testPERCENTILE()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testPERCENTILE(book);
	}
	
	// #NAME?
	@Test
	public void testSTDEVPA()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testSTDEVPA(book);
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIF()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testAVERAGEIF(book);
	}
	
	// #NAME?
	@Test
	public void testQUARTILE()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testQUARTILE(book);
	}
	
	// #NAME?
	@Test
	public void testMORMINV()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testMORMINV(book);
	}
	
	// #NAME?
	@Test
	public void testNEGBINOMDIST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testNEGBINOMDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCHITEST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testCHITEST(book);
	}
	
	// #NAME?
	@Test
	public void testBETADIST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testBETADIST(book);
	}
	
	// #NAME?
	@Test
	public void testVARA()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testVARA(book);
	}
	
	// #NAME?
	@Test
	public void testPROB()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		 testPROB(book);
	}
	
	// #NAME?
	@Test
	public void testZTEST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testZTEST(book);
	}
	
	// #NAME?
	@Test
	public void testVARPA()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testVARPA(book);
	}
	
	// #NAME?
	@Test
	public void testTTEST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testTTEST(book);
	}
	
	// #NAME?
	@Test
	public void testTREND()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testTREND(book);
	}
	
	// #NAME?
	@Test
	public void testSTEYX()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testSTEYX(book);
	}
	
	// #NAME?
	@Test
	public void testFTEST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testFTEST(book);
	}
	
	// #NAME?
	@Test
	public void testCOVAR()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testCOVAR(book);
	}
	
	// #NAME?
	@Test
	public void testLOGNORMDIST()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testLOGNORMDIST(book);
	}
	
	// #NAME?
	@Test
	public void testNORMSINV()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testNORMSINV(book);
	}
	
	// #NAME?
	@Test
	public void testSTANDARDIZE()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testSTANDARDIZE(book);
	}
	
	// #NAME?
	@Test
	public void testTRIMMEAN()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testTRIMMEAN(book);
	}
	
	// #NAME?
	@Test
	public void testCONFIDENCE()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testCONFIDENCE(book);
	}
	
	// #NAME?
	@Test
	public void testRSQ()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testRSQ(book);
	}
	
	// #NAME?
	@Test
	public void testBETAINV()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testBETAINV(book);
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIFS()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testAVERAGEIFS(book);
	}
	
	// #NAME?
	@Test
	public void testPEARSON()  {
		Book book = Util.loadBook(this,"TestFile2007-Formula.xlsx");
		testPEARSON(book);
	}
	
}
