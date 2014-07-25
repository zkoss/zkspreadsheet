package org.zkoss.zss.api.impl;

import java.io.IOException;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

@Ignore
public class RangeAPIIgnoredTest extends RangeAPITestBase {
	
	// Messagebox.show("A workbook must contain at least one visible worksheet");
	// API shouldn't bind to UI directly
	@Test
	public void testDeleteAllSheet2003() throws IOException {
		Book book = Util.loadBook("blank.xls");
		testDeleteAllSheet(book);
	}
	
	// Messagebox.show("A workbook must contain at least one visible worksheet");
	// API shouldn't bind to UI directly
	@Test
	public void testDeleteAllSheet2007() throws IOException {
		Book book = Util.loadBook("blank.xlsx");
		testDeleteAllSheet(book);
	}

}
