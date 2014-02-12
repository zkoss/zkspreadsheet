package org.zkoss.zss.api.impl.formula;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class DateTime2003Test extends FormulaTestBase {
	
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
	
	// DATE
	@Test
	public void testDATE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDATE(book);
	}

	// DATEVALUE
	@Test
	public void testDATEVALUE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDATEVALUE(book);
	}

	// DAY
	@Test
	public void testDAY() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDAY(book);
	}

	// DAY360
	@Test
	public void testDAY360() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testDAY360(book);
	}

	// HOUR
	@Test
	public void testHOUR() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testHOUR(book);
	}
	
	@Test
	public void testHOURWithTimeString()  {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testHOURWithString(book);
	}

	// MINUTE
	@Test
	public void testMINUTE() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMINUTE(book);
	}
	
	@Test
	public void testMINUTEWithTimeString()  {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMINUTEWithString(book);
	}

	// MONTH
	@Test
	public void testMONTH() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testMONTH(book);
	}

	// NETWORKDAYS
	@Test
	public void testNETWORKDAYS() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNETWORKDAYS(book);
	}
	
	@Test
	public void testNetworkDaysStartDateIsHoliday() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNetworkDaysStartDateIsHoliday(book);
	}

	@Test
	public void testEmptyDate() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testEmptyDate(book);
	}

	@Test
	public void testNetworkDaysAllHoliday() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNetworkDaysAllHoliday(book);
	}

	@Test
	public void testNetworkDaysSpecificHoliday() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNetworkDaysSpecificHoliday(book);
	}
	
	// SECOND
	@Test
	public void testSECOND() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSECOND(book);
	}
	
	@Test
	public void testSECONDWithTimeString()  {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testSECONDWithString(book);
	}

	// TIME
	@Test
	public void testTIME() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testTIME(book);
	}

	// WEEKDAY
	@Test
	public void testWEEKDAY() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testWEEKDAY(book);
	}
	
	
	// WORKDAY
	@Test
	public void testWORKDAY() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testWORKDAY(book);
	}

	@Test
	public void testStartDateIsHoliday() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testStartDateIsHoliday(book);
	}

	@Test
	public void testEndDateIsHoliday() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testEndDateIsHoliday(book);
	}

	@Test
	public void testNegativeWorkdayEndDateIsHolday() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testNegativeWorkdayEndDateIsHolday(book);
	}

	@Test
	public void testWorkdayBoundary() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testWorkdayBoundary(book);
	}

	@Test
	public void testWorkdaySpecifiedholiday() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testWorkdaySpecifiedholiday(book);
	}

	@Test
	public void testStartDateEqualsEndDate() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testStartDateEqualsEndDate(book);
	}

	@Test
	public void te8stStartDateLaterThanEndDate() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testStartDateLaterThanEndDate(book);
	}

	// YEAR
	@Test
	public void testYEAR() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testYEAR(book);
	}

	// YEARFRAC
	@Test
	public void testYEARFRAC() {
		Book book = Util.loadBook(this,"TestFile2003-Formula.xls");
		testYEARFRAC(book);
	}
	
}
