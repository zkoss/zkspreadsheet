package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class Engineering2007Test extends FormulaTestBase {
	
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
	
	// BESSELI
	@Test
	public void testBESSELI() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBESSELI(book);
	}

	// BESSELJ
	@Test
	public void testBESSELJ() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBESSELJ(book);
	}

	// BESSELK
	@Test
	public void testBESSELK() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBESSELK(book);
	}

	// BESSELY
	@Test
	public void testBESSELY() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBESSELY(book);
	}

	// BIN2DEC
	@Test
	public void testBIN2DEC() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBIN2DEC(book);
	}

	// BIN2HEX
	@Test
	public void testBIN2HEX() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBIN2HEX(book);
	}

	// BIN2OCT
	@Test
	public void testBIN2OCT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testBIN2OCT(book);
	}
	
	// COMPLEX
	@Test
	public void testCOMPLEX()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testCOMPLEX(book);
	}
	

	// DEC2BIN
	@Test
	public void testDEC2BIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testDEC2BIN(book);
	}

	// DEC2HEX
	@Test
	public void testDEC2HEX() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testDEC2HEX(book);
	}

	// DEC2OCT
	@Test
	public void testDEC2OCT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testDEC2OCT(book);
	}

	// DELTA
	@Test
	public void testDELTA() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testDELTA(book);
	}

	// ERF
	@Test
	public void testERF() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testERF(book);
	}

	// ERFC
	@Test
	public void testERFC() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testERFC(book);
	}

	// GESTEP
	@Test
	public void testGESTEP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testGESTEP(book);
	}

	// HEX2BIN
	@Test
	public void testHEX2BIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testHEX2BIN(book);
	}

	// HEX2DEC
	@Test
	public void testHEX2DEC() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testHEX2DEC(book);
	}

	// HEX2OCT
	@Test
	public void testHEX2OCT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testHEX2OCT(book);
	}

	// IMABS
	@Test
	public void testIMABS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMABS(book);
	}

	// IMAGINARY
	@Test
	public void testIMAGINARY() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMAGINARY(book);
	}

	// IMARGUMENT
	@Test
	public void testIMARGUMENT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMARGUMENT(book);
	}
	
	// IMCONJUGATE
	@Test
	public void testIMCONJUGATE() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMCONJUGATE(book);
	}
	
	// IMCOS
	@Test
	public void testIMCOS() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMCOS(book);
	}
	
	// IMDIV
	@Test
	public void testIMDIV() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMDIV(book);
	}
	
	// IMEXP
	@Test
	public void testIMEXP() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMEXP(book);
	}
	
	// IMLN
	@Test
	public void testIMLN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMLN(book);
	}

	// IMLOG10
	@Test
	public void testIMLOG10() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMLOG10(book);
	}
	
	// IMLOG2
	@Test
	public void testIMLOG2() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMLOG2(book);
	}

	// IMPOWER
	@Test
	public void testIMPOWER() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMPOWER(book);
	}
	
	// IMPODUCT
	@Test
	public void testIMPRODUCT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMPRODUCT(book);
	}
	
	// IMREAL
	@Test
	public void testIMREAL() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMREAL(book);
	}
	
	// IMSIN
	@Test
	public void testIMSIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMSIN(book);
	}

	// IMSQRT
	@Test
	public void testIMSQRT() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMSQRT(book);

	}
	
	// IMSUB
	@Test
	public void testIMSUB() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMSUB(book);
	}

	// IMSUM
	@Test
	public void testIMSUM() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testIMSUM(book);

	}

	// OCT2BIN
	@Test
	public void testOCT2BIN() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testOCT2BIN(book);
	}

	// OCT2DEC
	@Test
	public void testOCT2DEC() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testOCT2DEC(book);
	}

	// OCT2HEX
	@Test
	public void testOCT2HEX() {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testOCT2HEX(book);
	}

}
