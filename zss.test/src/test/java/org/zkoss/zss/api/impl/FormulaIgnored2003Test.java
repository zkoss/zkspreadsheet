package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author kuro, Hawk
 */
@Ignore
public class FormulaIgnored2003Test extends FormulaTestBase {
	
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

	/*
	// unsupported
	@Test
	public void testTRANSPOSE()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		Sheet sheet = book.getSheet("formula-lookup");
		// testTRANSPOSE(book);
	}
	
	// unsupported
	@Test
	public void testRTD()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		Sheet sheet = book.getSheet("formula-lookup");
		// testRTD(book);
	}
	*/
	
	@Test
	public void testAREAS()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testAREAS(book);
	}
	
	// slightly diffrent because of space
	@Test
	public void testINDEX()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testINDEX(book);
	}
	
	// 1990/1/1 is Monday, but Excel think it is not work day.
	@Test
	public void test19900101IsNotWorkDayInExcel()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		test19900101IsNotWorkDayInExcel(book);
	}
	
	// expected:<[1]> but was:<[2]>
	// different specification
	@Test
	public void testStartDateEmpty()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testStartDateEmpty(book);
	}
	
	// This should be check by human
	@Test
	public void testRANDBETWEEN()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testRANDBETWEEN(book);
	}
	
	// This should be check by human
	@Test
	public void testRAND()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testRAND(book);
	}
	
	// expected:<[0.03]> but was:<[-1.00]>
	@Test
	public void testGAMMADIST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testGAMMADIST(book);
	}
	
	// expected:<0.0[5]> but was:<0.0[3]>
	@Test
	public void testTDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTDIST(book);
	}
	
	// expected:<[15.2]1> but was:<[0.1]1>
	@Test
	public void testFINV()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFINV(book);
	}
	
	// expected:<[5.03]> but was:<[0.39]>
	@Test
	public void testHARMEAN()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testHARMEAN(book);
	}
	
	// expected:<1.[96]> but was:<1.[63]>
	@Test
	public void testTINV()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTINV(book);
	}
	
	// #NUM!
	@Test
	public void testSTDEVP()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testSTDEVP(book);
	}
	
	// expected:<[$1,234.5]7> but was:<[1234.56]7>
	@Test
	public void testDOLLAR()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testDOLLAR(book);
	}
	
	// #VALUE!
	@Test
	public void testHOURWithString()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testHOURWithString(book);
	}
	
	// #VALUE!
	@Test
	public void testVALUE()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testVALUE(book);
	}
	
	// #VALUE!
	@Test
	public void testEndDateEmpty()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testEndDateEmpty(book);
	}
	
	// #NAME?
	@Test
	public void testTOUSD2NTD()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTOUSD2NTD(book);
	}
	
	// #NAME?
	@Test
	public void testTOTWD()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTOTWD(book);
	}
	
	// #NAME?
	@Test
	public void testLOGEST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testLOGEST(book);
	}
	
	// #NAME?
	@Test
	public void testMIRR()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testMIRR(book);
	}
	
	// #NAME?
	@Test
	public void testISPMT()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-64814.81", Ranges.range(sheet, "B70").getCellData()
				.getFormatText());	
	}
	
	// #NAME?
	@Test
	public void testVDB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testVDB(book);
	}
	
	// #NAME?
	@Test
	public void testTIMEVALUE()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTIMEVALUE(book);
	}
	
	// #NAME?
	@Test
	public void testSERIESSUM()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testSERIESSUM(book);
	}
	
	// #NAME?
	@Test
	public void testCELL()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testCELL(book);
	}
	
	// #NAME?
	@Test
	public void testINFO()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testINFO(book);
	}
	
	// #NAME?
	@Test
	public void testASC()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testASC(book);
	}
	
	// #NAME?
	@Test
	public void testFINDB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFINDB(book);
	}
	
	// #NAME?
	@Test
	public void testLEFTB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testLEFTB(book);
	}
	
	// #NAME?
	@Test
	public void testLENB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testLENB(book);
	}
	
	// #NAME?
	@Test
	public void testMIDB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testMIDB(book);
	}
	
	// #NAME?
	@Test
	public void testSEARCHB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testSEARCHB(book);
	}
	
	// #NAME?
	@Test
	public void testREPLACEB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testREPLACEB(book);
	}
	
	// #NAME?
	@Test
	public void testRIGHTB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testRIGHTB(book);
	}
	
	// #NAME?
	@Test
	public void testFORECAST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFORECAST(book);
	}
	
	// #NAME?
	@Test
	public void testNORMSDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testNORMSDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCRITBINOM()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testCRITBINOM(book);
	}
	
	// #NAME?
	@Test
	public void testINTERCEPT()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testINTERCEPT(book);
	}
	
	// #NAME?
	@Test
	public void testFISHERINV()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFISHERINV(book);
	}
	
	// #NAME?
	@Test
	public void testSTDEVA()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testSTDEVA(book);
	}
	
	// #NAME?
	@Test
	public void testPERMUT()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testPERMUT(book);
	}
	
	// #NAME?
	@Test
	public void testLOGINV()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testLOGINV(book);
	}
	
	// #NAME?
	@Test
	public void testLINEST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testLINEST(book);
	}
	
	// #NAME?
	@Test
	public void testFREQUENCY()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFREQUENCY(book);
	}
	
	// #NAME?
	@Test
	public void testGROWTH()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testGROWTH(book);
	}
	
	// #NAME?
	@Test
	public void testFISHER()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFISHER(book);
	}
	
	// #NAME?
	@Test
	public void testPERCENTRANK()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testPERCENTRANK(book);
	}
	
	// #NAME?
	@Test
	public void testCORREL()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testCORREL(book);
	}
	
	// #NAME?
	@Test
	public void testPERCENTILE()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testPERCENTILE(book);
	}
	
	// #NAME?
	@Test
	public void testSTDEVPA()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testSTDEVPA(book);
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIF()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testAVERAGEIF(book);
	}
	
	// #NAME?
	@Test
	public void testQUARTILE()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testQUARTILE(book);
	}
	
	// #NAME?
	@Test
	public void testMORMINV()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testMORMINV(book);
	}
	
	// #NAME?
	@Test
	public void testNEGBINOMDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testNEGBINOMDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCHITEST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testCHITEST(book);
	}
	
	// #NAME?
	@Test
	public void testBETADIST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testBETADIST(book);
	}
	
	// #NAME?
	@Test
	public void testVARA()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testVARA(book);
	}
	
	// #NAME?
	@Test
	public void testPROB()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		 testPROB(book);
	}
	
	// #NAME?
	@Test
	public void testZTEST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testZTEST(book);
	}
	
	// #NAME?
	@Test
	public void testVARPA()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testVARPA(book);
	}
	
	// #NAME?
	@Test
	public void testTTEST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTTEST(book);
	}
	
	// #NAME?
	@Test
	public void testTREND()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTREND(book);
	}
	
	// #NAME?
	@Test
	public void testSTEYX()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testSTEYX(book);
	}
	
	// #NAME?
	@Test
	public void testFTEST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFTEST(book);
	}
	
	// #NAME?
	@Test
	public void testFDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testFDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCOVAR()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testCOVAR(book);
	}
	
	// #NAME?
	@Test
	public void testLOGNORMDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testLOGNORMDIST(book);
	}
	
	// #NAME?
	@Test
	public void testNORMSINV()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testNORMSINV(book);
	}
	
	// #NAME?
	@Test
	public void testSTANDARDIZE()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testSTANDARDIZE(book);
	}
	
	// #NAME?
	@Test
	public void testTRIMMEAN()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testTRIMMEAN(book);
	}
	
	// #NAME?
	@Test
	public void testCONFIDENCE()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testCONFIDENCE(book);
	}
	
	// #NAME?
	@Test
	public void testRSQ()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testRSQ(book);
	}
	
	// #NAME?
	@Test
	public void testBETAINV()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testBETAINV(book);
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIFS()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testAVERAGEIFS(book);
	}
	
	// #NAME?
	@Test
	public void testPEARSON()  {
		Book book = Util.loadBook(this,"book/TestFile2003-Formula.xls");
		testPEARSON(book);
	}
	
}
