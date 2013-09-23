package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author kuro, Hawk
 */
@Ignore
public class FormulaIgnored2007Test extends FormulaTestBase {
	
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
	public void testTRANSPOSE() throws IOException {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
	}
	
	// unsupported
	@Test
	public void testRTD() throws IOException {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-lookup");
	}
	*/
	
	@Test
	public void testAREAS()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testAREAS(book);
	}
	
	// 1990/1/1 is Monday, but Excel think it is not work day.
	@Test
	public void test19900101IsNotWorkDayInExcel()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		test19900101IsNotWorkDayInExcel(book);
	}
	
	
	// slightly different because of space
	@Test
	public void testCOMPLEX()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testCOMPLEX(book);
	}
	
	// slightly different because of space
	@Test
	public void testINDEX()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testINDEX(book);
	}
	
	// expected:<[1]> but was:<[2]>
	// different specification
	@Test
	public void testStartDateEmpty()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testStartDateEmpty(book);
	}
	
	// This should be check by human
	@Test
	public void testRANDBETWEEN()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testRANDBETWEEN(book);
	}
	
	// This should be check by human
	@Test
	public void testRAND()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testRAND(book);
	}
	
	// expected:<[0.03]> but was:<[-1.00]>
	@Test
	public void testGAMMADIST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testGAMMADIST(book);
	}
	
	// expected:<0.0[5]> but was:<0.0[3]>
	@Test
	public void testTDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTDIST(book);
	}
	
	// expected:<[15.2]1> but was:<[0.1]1>
	@Test
	public void testFINV()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFINV(book);
	}
	
	// expected:<[5.03]> but was:<[0.39]>
	@Test
	public void testHARMEAN()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testHARMEAN(book);
	}
	
	// expected:<1.[96]> but was:<1.[63]>
	@Test
	public void testTINV()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTINV(book);
	}
	
	// #NUM!
	@Test
	public void testSTDEVP()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testSTDEVP(book);
	}
	
	// expected:<[$1,234.5]7> but was:<[1234.56]7>
	@Test
	public void testDOLLAR()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testDOLLAR(book);
	}
	
	// #VALUE!
	@Test
	public void testHOURWithString()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testHOURWithString(book);
	}
	
	// #VALUE!
	@Test
	public void testVALUE()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testVALUE(book);
	}
	
	// #VALUE!
	@Test
	public void testEndDateEmpty()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testEndDateEmpty(book);
	}
	
	// #NAME?
	@Test
	public void testTOUSD2NTD()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTOUSD2NTD(book);
	}
	
	// #NAME?
	@Test
	public void testTOTWD()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTOTWD(book);
	}
	
	// #NAME?
	@Test
	public void testLOGEST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testLOGEST(book);
	}
	
	// #NAME?
	@Test
	public void testMIRR()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testMIRR(book);
	}
	
	// #NAME?
	@Test
	public void testISPMT()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		Sheet sheet = book.getSheet("formula-financial");
		assertEquals("-64814.81", Ranges.range(sheet, "B70").getCellData()
				.getFormatText());	
	}
	
	// #NAME?
	@Test
	public void testVDB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testVDB(book);
	}
	
	// #NAME?
	@Test
	public void testTIMEVALUE()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTIMEVALUE(book);
	}
	
	// #NAME?
	@Test
	public void testSERIESSUM()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testSERIESSUM(book);
	}
	
	// #NAME?
	@Test
	public void testCELL()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testCELL(book);
	}
	
	// #NAME?
	@Test
	public void testINFO()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testINFO(book);
	}
	
	// #NAME?
	@Test
	public void testASC()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testASC(book);
	}
	
	// #NAME?
	@Test
	public void testFINDB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFINDB(book);
	}
	
	// #NAME?
	@Test
	public void testLEFTB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testLEFTB(book);
	}
	
	// #NAME?
	@Test
	public void testLENB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testLENB(book);
	}
	
	// #NAME?
	@Test
	public void testMIDB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testMIDB(book);
	}
	
	// #NAME?
	@Test
	public void testSEARCHB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testSEARCHB(book);
	}
	
	// #NAME?
	@Test
	public void testREPLACEB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testREPLACEB(book);
	}
	
	// #NAME?
	@Test
	public void testRIGHTB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testRIGHTB(book);
	}
	
	// #NAME?
	@Test
	public void testFORECAST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFORECAST(book);
	}
	
	// #NAME?
	@Test
	public void testNORMSDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testNORMSDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCRITBINOM()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testCRITBINOM(book);
	}
	
	// #NAME?
	@Test
	public void testINTERCEPT()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testINTERCEPT(book);
	}
	
	// #NAME?
	@Test
	public void testFISHERINV()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFISHERINV(book);
	}
	
	// #NAME?
	@Test
	public void testSTDEVA()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testSTDEVA(book);
	}
	
	// #NAME?
	@Test
	public void testPERMUT()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testPERMUT(book);
	}
	
	// #NAME?
	@Test
	public void testLOGINV()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testLOGINV(book);
	}
	
	// #NAME?
	@Test
	public void testLINEST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testLINEST(book);
	}
	
	// #NAME?
	@Test
	public void testFREQUENCY()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFREQUENCY(book);
	}
	
	// #NAME?
	@Test
	public void testGROWTH()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testGROWTH(book);
	}
	
	// #NAME?
	@Test
	public void testFISHER()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFISHER(book);
	}
	
	// #NAME?
	@Test
	public void testPERCENTRANK()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testPERCENTRANK(book);
	}
	
	// #NAME?
	@Test
	public void testCORREL()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testCORREL(book);
	}
	
	// #NAME?
	@Test
	public void testPERCENTILE()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testPERCENTILE(book);
	}
	
	// #NAME?
	@Test
	public void testSTDEVPA()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testSTDEVPA(book);
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIF()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testAVERAGEIF(book);
	}
	
	// #NAME?
	@Test
	public void testQUARTILE()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testQUARTILE(book);
	}
	
	// #NAME?
	@Test
	public void testMORMINV()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testMORMINV(book);
	}
	
	// #NAME?
	@Test
	public void testNEGBINOMDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testNEGBINOMDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCHITEST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testCHITEST(book);
	}
	
	// #NAME?
	@Test
	public void testBETADIST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testBETADIST(book);
	}
	
	// #NAME?
	@Test
	public void testVARA()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testVARA(book);
	}
	
	// #NAME?
	@Test
	public void testPROB()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		 testPROB(book);
	}
	
	// #NAME?
	@Test
	public void testZTEST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testZTEST(book);
	}
	
	// #NAME?
	@Test
	public void testVARPA()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testVARPA(book);
	}
	
	// #NAME?
	@Test
	public void testTTEST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTTEST(book);
	}
	
	// #NAME?
	@Test
	public void testTREND()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTREND(book);
	}
	
	// #NAME?
	@Test
	public void testSTEYX()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testSTEYX(book);
	}
	
	// #NAME?
	@Test
	public void testFTEST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFTEST(book);
	}
	
	// #NAME?
	@Test
	public void testFDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testFDIST(book);
	}
	
	// #NAME?
	@Test
	public void testCOVAR()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testCOVAR(book);
	}
	
	// #NAME?
	@Test
	public void testLOGNORMDIST()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testLOGNORMDIST(book);
	}
	
	// #NAME?
	@Test
	public void testNORMSINV()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testNORMSINV(book);
	}
	
	// #NAME?
	@Test
	public void testSTANDARDIZE()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testSTANDARDIZE(book);
	}
	
	// #NAME?
	@Test
	public void testTRIMMEAN()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testTRIMMEAN(book);
	}
	
	// #NAME?
	@Test
	public void testCONFIDENCE()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testCONFIDENCE(book);
	}
	
	// #NAME?
	@Test
	public void testRSQ()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testRSQ(book);
	}
	
	// #NAME?
	@Test
	public void testBETAINV()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testBETAINV(book);
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIFS()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testAVERAGEIFS(book);
	}
	
	// #NAME?
	@Test
	public void testPEARSON()  {
		Book book = Util.loadBook(this,"book/TestFile2007-Formula.xlsx");
		testPEARSON(book);
	}
	
}
