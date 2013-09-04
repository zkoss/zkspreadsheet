package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * ZSS-36
 * @author kuro
 */
public class IssueTest {
	private Book _workbook;
	
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
	
	/**
	 * Exception when exporting excel twice.
	 */
	@Test
	public void testZSS36() throws IOException {
		final String filename = "book/blank.xlsx";
		final InputStream is = IssueTest.class.getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		export();
		export();
	}
	
	private void export() throws IOException {
		Exporter excelExporter = Exporters.getExporter("excel");
		FileOutputStream fos = new FileOutputStream(new File(ShiftTest.class.getResource("").getPath() + "book/test.xlsx"));
		excelExporter.export(_workbook, fos);
	}
}
