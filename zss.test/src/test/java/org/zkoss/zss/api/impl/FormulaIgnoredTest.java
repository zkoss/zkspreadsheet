package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

@Ignore
public class FormulaIgnoredTest {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN,-1));
	}
	
	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}
	
	// expected:<[0.03]> but was:<[-1.00]>
	@Test
	public void testGAMMADIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.03", Ranges.range(sheet, "B76").getCellFormatText());
	}
	
	// expected:<0.0[5]> but was:<0.0[3]>
	@Test
	public void testTDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.05", Ranges.range(sheet, "B175").getCellFormatText());
	}
	
	// expected:<[15.2]1> but was:<[0.1]1>
	@Test
	public void testFINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("15.21", Ranges.range(sheet, "B61").getCellFormatText());
	}
	
	// expected:<[5.03]> but was:<[0.39]>
	@Test
	public void testHARMEAN() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.03", Ranges.range(sheet, "B89").getCellFormatText());
	}
	
	// expected:<1.[96]> but was:<1.[63]>
	@Test
	public void testTINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.96", Ranges.range(sheet, "B177").getCellFormatText());
	}
	
	// #NUM!
	@Test
	public void testSTDEVP() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("26.05", Ranges.range(sheet, "B168").getCellFormatText());
	}
	
	// expected:<[$1,234.5]7> but was:<[1234.56]7>
	@Test
	public void testDOLLAR() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("$1,234.57", Ranges.range(sheet, "B14").getCellFormatText());
	}
	
	// #VALUE!
	@Test
	public void testVALUE() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("0.7", Ranges.range(sheet, "B73").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testCELL() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("3", Ranges.range(sheet, "B3").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testINFO() throws IOException {
		Book book = Util.loadBook("book/266-info-formula.xlsx");
		Sheet sheet = book.getSheet("formula-info");
		assertEquals("12.0", Ranges.range(sheet, "B8").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testASC() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("EXCEL", Ranges.range(sheet, "B3").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testFINDB() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("1", Ranges.range(sheet, "B22").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testLEFTB() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("\u6771", Ranges.range(sheet, "B29").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testLENB() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("11", Ranges.range(sheet, "B33").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testMIDB() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Fluid", Ranges.range(sheet, "B40").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testSEARCHB() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("6", Ranges.range(sheet, "B57").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testREPLACEB() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("abcde*k", Ranges.range(sheet, "B47").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testRIGHTB() throws IOException {
		Book book = Util.loadBook("book/264-text-formula.xlsx");
		Sheet sheet = book.getSheet("formula-text");
		assertEquals("Price", Ranges.range(sheet, "B53").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testFORECAST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("10.61", Ranges.range(sheet, "B67").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testNORMSDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.91", Ranges.range(sheet, "B129").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testCRITBINOM() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("4", Ranges.range(sheet, "B53").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testINTERCEPT() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.05", Ranges.range(sheet, "B94").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testFISHERINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.75", Ranges.range(sheet, "B65").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testSTDEVA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("27.46", Ranges.range(sheet, "B166").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testPERMUT() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("970200", Ranges.range(sheet, "B140").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testLOGINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("4.00", Ranges.range(sheet, "B106").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testLINEST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B103").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testFREQUENCY() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1", Ranges.range(sheet, "B70").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testGROWTH() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("320196.72", Ranges.range(sheet, "B85").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testFISHER() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.97", Ranges.range(sheet, "B63").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testPERCENTRANK() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.33", Ranges.range(sheet, "B138").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testCORREL() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.9971", Ranges.range(sheet, "B37").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testPERCENTILE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.9", Ranges.range(sheet, "B136").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testSTDEVPA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("26.05", Ranges.range(sheet, "B170").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIF() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("14000", Ranges.range(sheet, "B11").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testQUARTILE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3.5", Ranges.range(sheet, "B147").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testMORMINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("42.00", Ranges.range(sheet, "B127").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testNEGBINOMDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.06", Ranges.range(sheet, "B123").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testCHITEST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.000308", Ranges.range(sheet, "B27").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testBETADIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.69", Ranges.range(sheet, "B17").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testVARA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("754.27", Ranges.range(sheet, "B191").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testPROB() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.1", Ranges.range(sheet, "B144").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testZTEST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.09", Ranges.range(sheet, "B199").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testVARPA() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("678.84", Ranges.range(sheet, "B195").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testTTEST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.20", Ranges.range(sheet, "B186").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testTREND() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("133953.33", Ranges.range(sheet, "B179").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testSTEYX() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3.31", Ranges.range(sheet, "B172").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testFTEST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.65", Ranges.range(sheet, "B73").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testFDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.01", Ranges.range(sheet, "B59").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testCOVAR() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("5.2", Ranges.range(sheet, "B50").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testLOGNORMDIST() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.04", Ranges.range(sheet, "B108").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testNORMSINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.33", Ranges.range(sheet, "B131").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testSTANDARDIZE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("1.33", Ranges.range(sheet, "B162").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testTRIMMEAN() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("3.78", Ranges.range(sheet, "B183").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testCONFIDENCE() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.69", Ranges.range(sheet, "B35").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testRSQ() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.061", Ranges.range(sheet, "B151").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testBETAINV() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("2", Ranges.range(sheet, "B19").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testAVERAGEIFS() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("87.5", Ranges.range(sheet, "B14").getCellFormatText());
	}
	
	// #NAME?
	@Test
	public void testPEARSON() throws IOException {
		Book book = Util.loadBook("book/270-statistical.xlsx");
		Sheet sheet = book.getSheet("formula-statistical");

		assertEquals("0.70", Ranges.range(sheet, "B133").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMCOS() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.833730025131149 - 0.988897705762865i", Ranges.range(sheet, "B47").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMLN() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1.6094379124341 + 0.927295218001612i", Ranges.range(sheet, "B53").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMSQRT() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("1.09868411346781 + 0.455089860562227i", Ranges.range(sheet, "B67").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMLOG10() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("0.698970004336019+0.402719196273373i", Ranges.range(sheet, "B55").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMSUM() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("8 + i", Ranges.range(sheet, "B71").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMSIN() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("3.85373803791938 - 27.0168132580039i", Ranges.range(sheet, "B65").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMSUB() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("8 + i", Ranges.range(sheet, "B69").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMLOG2() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("2.32192809506607+1.33780421255394i", Ranges.range(sheet, "B57").getCellFormatText());
	}
	
	// Little different
	@Test
	public void testIMPOWER() throws IOException {
		Book book = Util.loadBook("book/271-engineering.xlsx");
		Sheet sheet = book.getSheet("formula-engineering");
		assertEquals("-46 + 9.00000000000001i", Ranges.range(sheet, "B59").getCellFormatText());
	}
}
