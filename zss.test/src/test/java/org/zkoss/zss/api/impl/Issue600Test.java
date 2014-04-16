package org.zkoss.zss.api.impl;

import static org.junit.Assert.*;

import java.io.*;
import java.util.Locale;

import org.junit.*;
import org.zkoss.zss.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.model.InvalidModelOpException;

/**
 * @author Hawk
 *
 */
public class Issue600Test {
	
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
	
	
	@Test
	public void testZSS610(){
		Book book;
		book = Util.loadBook(this,"book/blank.xlsx");
		Assert.assertEquals(Book.BookType.XLSX,book.getType());
		book = Util.loadBook(this,"book/blank.xls");
		Assert.assertEquals(Book.BookType.XLS,book.getType());
		
		
	}
	
	@Test
	public void testZSS621() throws Exception {
		testZSS621("book/621-shared-formula.xlsx");
		testZSS621("book/621-shared-formula.xls");
	}

	public void testZSS621(String path) throws Exception {
		String[] headers = { "C3", "B6", "C13", "B14", "C14", "D14", "E14",
				"F14", "G14", "H14" };
		Double[] headerValues = { 2.0, 8.0, 2.0, 8.0, 9.0, 10.0, 11.0, 12.0,
				13.0, 14.0 };

		// follow issue description
		// 1. import a book
		Book book = Util.loadBook(this, path);
		Sheet sheet = book.getSheetAt(0);
		for (int i = 0; i < headers.length; ++i) {
			assertEquals(headers[i], headerValues[i],
					Ranges.range(sheet, headers[i]).getCellData().getDoubleValue());
		}
		
		// 2. clear cell with shared formula header
		for (String header : headers) {
			Ranges.range(sheet, header).clearContents();
		}
		for (int i = 0; i < headers.length; ++i) {
			assertTrue(headers[i], Ranges.range(sheet, headers[i]).getCellData().isBlank());
		}
		
		// 3. export then import again
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		Exporters.getExporter().export(book, out);
		out.close();
		byte[] data = out.toByteArray();
		Book book2 = Importers.getImporter("excel").imports(new ByteArrayInputStream(data), "exported");
		
		// 4. read data again
		Sheet sheet2 = book2.getSheetAt(0);
		Ranges.range(sheet2, "L1:R7").clearContents();
		for (int i = 0; i < headers.length; ++i) {
			assertTrue(headers[i], Ranges.range(sheet2, headers[i]).getCellData().isBlank());
		}
		String[] areas = {"D3:H3", "B7:B11", "D13:H13", "B15:B19", "C15:C19", "D15:D19", "E15:E19", "F15:F19", "G15:G19", "H15:H19"};
		for (String area : areas) {
			AreaRef af = new AreaRef(area);
			for (int r = af.getRow(); r <= af.getLastRow(); ++r) {
				for (int c = af.getColumn(); c <= af.getLastColumn(); ++c) {
					assertEquals(new Double(0.0), Ranges.range(sheet2, r, c).getCellValue());
				}
			}
		}
	}
	
	@Test
	public void testZSS624() {
		Book book = Util.loadBook(this, "book/blank.xlsx");
		Sheet sheet = book.getSheetAt(0);
		Ranges.range(sheet, "B1").setCellEditText("=SUM(1+2+3)");
		assertEquals("=SUM(1+2+3)", Ranges.range(sheet, "B1").getCellEditText());
		
		Ranges.range(sheet, "B1").paste(Ranges.range(sheet, "B1"));
		Ranges.range(sheet, "A1:B1").paste(Ranges.range(sheet, "A1:B1"));
		Ranges.range(sheet, "A1:B1").paste(Ranges.range(sheet, "A1:D1"));
		
		Ranges.range(sheet, "B1").pasteSpecial(Ranges.range(sheet, "B1"), PasteType.VALUES, PasteOperation.NONE, false, false);
		assertEquals("6", Ranges.range(sheet, "B1").getCellEditText());
	}
	
	@Test
	public void testZSS660InvalidNamedRange(){
		Book book = Util.loadBook(this, "book/blank.xlsx");
		Sheet sheet = book.getSheetAt(0);
		String invalidNameList[] = {"1A", "123", "A1", "c", "mya1", "have space", buildStringByLength(256)};
		int invalidCount = 0;
		for (String invalidName : invalidNameList){
			try{
				Ranges.range(sheet, "A1").createName(invalidName);
				fail(invalidName+" shall not be a valid name for a named range.");
			}catch (InvalidModelOpException e) {
				invalidCount++;
			}
		}
		
		assertEquals(invalidCount, invalidNameList.length);
	}
	
	private String buildStringByLength(int length){
		StringBuilder name = new StringBuilder();
		for (int i =0 ; i <length ; i++){
			name.append("a");
		}
		return name.toString();
	}
	
	@Test
	public void testZSS660NamedRange(){
		Book book = Util.loadBook(this, "book/blank.xlsx");
		Sheet sheet = book.getSheetAt(0);
		String nameList[] = {"_a1", "\\a1", "a.b", "中文", "myname", buildStringByLength(255)};
		for (String validName : nameList){
			Ranges.range(sheet, "A1").createName(validName);
		}
		assertEquals(nameList.length, book.getInternalBook().getNames().size());
		try{
			Ranges.range(sheet, "A1").createName("MYNAME");
			fail("Should not create a duplicated named range: MYNAME ");
		}catch (InvalidModelOpException e) {
			//catch to avoid test failure
		}
	}
}
