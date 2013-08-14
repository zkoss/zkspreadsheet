package org.zkoss.zss.api.impl;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * style test
 * ZSS-290, after merging and unmerging cells, input and style both work incorrect
 * 
 * @author kuro
 *
 */
public class StyleTest {

	private static Book _workbook;
	
	@Before
	public void setUp() throws Exception {
		Setup.touch();
//		final String filename = "book/overlappingPaste.xlsx";
//		final InputStream is = PasteTest.class.getResourceAssStream(filename);
//		_workbook = Importers.getImporter().imports(is, filename);
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	@Test 
	public void testZSS290() {
//		Sheet sheet1 = _workbook.getSheet("Sheet1");
	}
}
