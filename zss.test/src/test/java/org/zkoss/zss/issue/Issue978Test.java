package org.zkoss.zss.issue;

import java.io.IOException;
import java.util.List;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.Validation;

public class Issue978Test {
	
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

	/**
	 * parse a defined name which contains question mark '?'
	 * @throws IOException 
	 */
	@Test
	public void conditionalSumproduct() throws IOException {
		Book book = Util.loadBook(this, "book/978-validation.xlsx");
		Sheet sheet1 = book.getSheetAt(0);
		Range rng = Ranges.range(sheet1, "A1");
		List<Validation> dvs = rng.getValidations();
		Assert.assertEquals("dvs.length", 1, dvs.size());
		Validation dv = dvs.get(0);
		
		Assert.assertEquals("dv.getValidationType()", Validation.ValidationType.INTEGER, dv.getValidationType());
		Assert.assertEquals("dv.getFormula1()", "1", dv.getFormula1());
		Assert.assertEquals("dv.getFormula2()", "10", dv.getFormula2());
	}
}