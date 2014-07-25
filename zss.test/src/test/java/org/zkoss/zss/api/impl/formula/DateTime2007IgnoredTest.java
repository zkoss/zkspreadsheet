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
public class DateTime2007IgnoredTest extends FormulaTestBase {
	
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
	
	// NETWORKDAYS
	// 1990/1/1 is Monday, but Excel think it is not work day.
	@Test
	public void test19900101IsNotWorkDayInExcel()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		test19900101IsNotWorkDayInExcel(book); // =NETWORKDAYS(DATE(1900,1,1),DATE(1900,1,1))
	}
	
	// expected:<[1]> but was:<[2]>
	// different specification
	@Test
	public void testStartDateEmpty()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testStartDateEmpty(book); // =NETWORKDAYS(B30,C30)
		// B30 : blank
		// C30 : 1990/1/2
	}
	
	// #VALUE!
	@Test
	public void testEndDateEmpty()  {
		Book book = Util.loadBook("TestFile2007-Formula.xlsx");
		testEndDateEmpty(book); // =NETWORKDAYS(C30, B30)
		// B30 : blank
		// C30 : 1990/1/2
	}
	
}
