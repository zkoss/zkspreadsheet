package org.zkoss.zss.api.impl;

import static org.junit.Assert.*;

import java.io.*;
import java.util.Locale;
import java.util.Set;

import org.junit.*;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.zss.*;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SName;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.range.SExporter;
import org.zkoss.zss.range.impl.imexp.ExcelXlsExporter;
import org.zkoss.zss.range.impl.imexp.ExcelXlsxExporter;
import org.zkoss.zss.range.impl.imexp.ExcelXlsxImporter;

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
	public void testZSS649ModifyName() {
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		book.createSheet("Sheet0");
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("B1").setValue(2);
		sheet1.getCell("A2").setValue(3);
		sheet1.getCell("B2").setValue(4);
		
		SName name = book.createName("FOO");
		name.setRefersToFormula("Sheet1!A1:B2");
		
		sheet1.getCell("C1").setValue("=SUM(FOO)");
		
		Assert.assertEquals(10D, sheet1.getCell("C1").getValue());
		
		// test shrink
		sheet1.deleteColumn(1, 1); // delete column B
		Assert.assertEquals("Sheet1!A1:A2", name.getRefersToFormula());
		Assert.assertEquals(4D, sheet1.getCell("B1").getValue());
		
		// test extend
		sheet1.insertColumn(1, 1); // insert column B
		Assert.assertEquals("Sheet1!A1:A2", name.getRefersToFormula());
		Assert.assertEquals(4D, sheet1.getCell("C1").getValue());
		
		// test extend
		sheet1.insertColumn(0, 0); // insert column A
		Assert.assertEquals("Sheet1!B1:B2", name.getRefersToFormula());
		Assert.assertEquals(4D, sheet1.getCell("D1").getValue());

		// test shrink
		sheet1.deleteColumn(0, 0); // delete column A
		Assert.assertEquals("Sheet1!A1:A2", name.getRefersToFormula());
		Assert.assertEquals(4D, sheet1.getCell("C1").getValue());

		// test shrink
		sheet1.deleteRow(1,  1); //  delete row 2
		Assert.assertEquals("Sheet1!A1:A1", name.getRefersToFormula());
		Assert.assertEquals(1D, sheet1.getCell("C1").getValue());

		// test extend
		sheet1.insertRow(0,  0); // insert row 1
		Assert.assertEquals("Sheet1!A2:A2", name.getRefersToFormula());
		Assert.assertEquals(1D, sheet1.getCell("C2").getValue());

		// test rename sheet
		book.setSheetName(sheet1, "SheetX");
		Assert.assertEquals("SheetX!A2:A2", name.getRefersToFormula());
		Assert.assertEquals(1D, sheet1.getCell("C2").getValue());
		
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("B1").setValue(2);
		sheet1.getCell("A2").setValue(3);
		sheet1.getCell("B2").setValue(4);
		
		name.setRefersToFormula("SheetX!A1:B2");
		Assert.assertEquals("SheetX!A1:B2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("C2").getValue());
		
		// test extend
		sheet1.insertColumn(1, 1); // insert column B
		Assert.assertEquals("SheetX!A1:C2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("D2").getValue());
		
		// test extend
		sheet1.insertColumn(0, 0); // insert column A
		Assert.assertEquals("SheetX!B1:D2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("E2").getValue());

		// test shrink
		sheet1.deleteColumn(1, 1); // delete column B
		Assert.assertEquals("SheetX!B1:C2", name.getRefersToFormula());
		Assert.assertEquals(6D, sheet1.getCell("D2").getValue());
		
		// test move 
		sheet1.moveCell(0, 0, 2, 2, 2, 4); // A1:C3 -> E3:G5
		Assert.assertEquals("SheetX!F3:G4", name.getRefersToFormula());
		Assert.assertEquals(6D, sheet1.getCell("D2").getValue());
		
		// test move
		sheet1.moveCell(2, 6, 3, 6, 0, 1); // G3:G4 -> H3:H4
		Assert.assertEquals("SheetX!F3:H4", name.getRefersToFormula());
		Assert.assertEquals(6D, sheet1.getCell("D2").getValue());

		// test move
		sheet1.moveCell(2, 5, 2, 7, -1, 0); // F3:H3 -> F2:H2
		Assert.assertEquals("SheetX!F2:H4", name.getRefersToFormula());
		Assert.assertEquals(6D, sheet1.getCell("D2").getValue());
		
		// test move
		sheet1.moveCell(2, 6, 2, 8, 1, 0); // F3:H3 -> F2:H2
		Assert.assertEquals("SheetX!F2:H4", name.getRefersToFormula());
		Assert.assertEquals(2D, sheet1.getCell("D2").getValue());

		// test insert/delete cells
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("B1").setValue(2);
		sheet1.getCell("A2").setValue(3);
		sheet1.getCell("B2").setValue(4);
		
		name.setRefersToFormula("SheetX!A1:B2");
		Assert.assertEquals("SheetX!A1:B2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("D2").getValue());
		
		// insert B1:B2
		sheet1.insertCell(0, 1, 1, 1, true); // A1:C2
		Assert.assertEquals("SheetX!A1:C2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("E2").getValue());
		
		// delete B1:B2
		sheet1.deleteCell(0, 1, 1, 1, true); // A1:B2
		Assert.assertEquals("SheetX!A1:B2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("D2").getValue());
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
	
	@Test 
	public void testZSS661ModifyName() {
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		book.createSheet("Sheet0");
		SSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("B1").setValue(2);
		sheet1.getCell("A2").setValue(3);
		sheet1.getCell("B2").setValue(4);
		
		SName name = book.createName("FOO");
		name.setRefersToFormula("Sheet1!A1:B1");
		
		sheet1.getCell("C1").setValue("=SUM(FOO)");
		
		Assert.assertEquals(3D, sheet1.getCell("C1").getValue());
		
		name.setRefersToFormula("Sheet1!A2:B2");
		Assert.assertEquals(7D, sheet1.getCell("C1").getValue());//shouldn't fail
		
		sheet1.getCell("A2").setValue(5);
		sheet1.getCell("B2").setValue(6);
		Assert.assertEquals(11D, sheet1.getCell("C1").getValue());
		
		book.setNameName(name, "BAR"); // formula that use FOO should be changed accordingly
		Assert.assertEquals(11D, sheet1.getCell("C1").getValue());
		
		sheet1.getCell("A2").setValue(7);
		sheet1.getCell("B2").setValue(8);
		Assert.assertEquals(15D, sheet1.getCell("C1").getValue());		

		sheet1.getCell("C1").setValue("=SUM(FOO)");
		Assert.assertEquals("#NAME?", sheet1.getCell("C1").getErrorValue().getErrorString());//shouldn't fail
	}
	
	@Test
	public void testZSS685ImportStyle(){
		Object[] books = _loadBooks(this, "book/685-StyleOverflow.xlsx");
		Book book = (Book) books[0];
		Workbook wbookimp = (Workbook) books[1];
		assertEquals(15982, wbookimp.getNumCellStyles());
		SBook sbook = book.getInternalBook();
		XlsxExporter exp = new XlsxExporter();
		try {
			exp.export(sbook, new ByteArrayOutputStream());
			Workbook wbook = exp.getWorkbook();
			assertEquals(268, wbook.getNumCellStyles());
		} catch (IOException e) {
			fail("Should export to byte array output stream without issue:\n");
			e.printStackTrace();
		}
	}
	
	@Test
	public void testZSS687ReferName() {
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("SheetX");

		SName name = book.createName("FOO");
		name.setRefersToFormula("SheetX!A1:B2");
		sheet1.getCell("D2").setValue("=SUM(FOO)");

		// test insert/delete cells
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("B1").setValue(2);
		sheet1.getCell("A2").setValue(3);
		sheet1.getCell("B2").setValue(4);

		Assert.assertEquals("SheetX!A1:B2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("D2").getValue());

		// insert B1:B1
		sheet1.insertCell(0, 1, 0, 1, true);
		Assert.assertEquals(null, sheet1.getCell("B1").getValue());
		Assert.assertEquals("SheetX!A1:B2", name.getRefersToFormula());
		Assert.assertEquals(8D, sheet1.getCell("D2").getValue());		

		// delete B1:B1
		sheet1.deleteCell(0, 1, 0, 1, true);
		Assert.assertEquals(2D, sheet1.getCell("B1").getValue());
		Assert.assertEquals("SheetX!A1:B2", name.getRefersToFormula());
		Assert.assertEquals(10D, sheet1.getCell("D2").getValue());
		
		// move B1:B1
		sheet1.moveCell(0, 1, 0, 1, 1, 1);
		Assert.assertEquals(2D, sheet1.getCell("C2").getValue());
		Assert.assertEquals("SheetX!A1:B2", name.getRefersToFormula());
		Assert.assertEquals(8D, sheet1.getCell("D2").getValue());
	}

	@Test
	public void testZSS694CopyValidation() {
		Object[] books = _loadBooks(this, "book/694-copy-validation.xlsx");
		Book book = (Book) books[0];
		Sheet sheet = book.getSheet("Sheet1");
		
		// 1. Validation_1 in A1 and A3:B4
		SSheet ssheet = sheet.getInternalSheet();
		SDataValidation a1dv = ssheet.getDataValidation(0, 0); // validation in  a1
		Set<CellRegion> a1regions = a1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_1", 2, a1regions.size());
		Assert.assertTrue("1st Validation_1 in A1", a1regions.contains(new CellRegion("a1"))); // a1
		Assert.assertTrue("1st Validation_1 in A3:B4", a1regions.contains(new CellRegion("a3:b4"))); // a3:b4
		
		// 2. Validation_2 in D1:E2
		SDataValidation d1dv = ssheet.getDataValidation(0, 3); // validation in  d1
		Set<CellRegion> d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions Validation_2", 1, d1regions.size());
		Assert.assertTrue("Validation_2 in D1:E2", d1regions.contains(new CellRegion("d1:e2"))); // d1:e2
		
		// 3. Cut A1 and paste to D1
		Range a1 = Ranges.range(sheet, "a1");
		a1.paste(Ranges.range(sheet, "d1"), true);
		
		// 3.1. Validation_1 no longer contains cell A1 but contains D1 and A3:B4 
		a1regions = a1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_1 after cut A1 to D1", 2, a1regions.size());
		Assert.assertFalse("Validation_1 not in A1 after cut A1 to D1", a1regions.contains(new CellRegion("a1"))); // a1
		Assert.assertTrue("Validation_1 in D1 after cut A1 to D1", a1regions.contains(new CellRegion("d1"))); // d1
		Assert.assertTrue("Validation_1 in A3:B4 after cut A1 to D1", a1regions.contains(new CellRegion("a3:b4"))); // a3:b4
		
		// 3.2. Validation_2 no longer contains cell D1 but split to two regions: E1 and D2:E2 
		d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_2", 2, d1regions.size());
		Assert.assertFalse("Validation_2 in D1", d1regions.contains(new CellRegion("d1"))); // d1
		Assert.assertTrue("Validation_2 in E1", d1regions.contains(new CellRegion("e1"))); // e1
		Assert.assertTrue("Validation_2 in D2:E2", d1regions.contains(new CellRegion("d2:e2"))); // d2:e2
		
		// 4. Cut B4 and paste to E2
		Range b4 = Ranges.range(sheet, "b4");
		b4.paste(Ranges.range(sheet, "e2"), true);
		
		// 4.1. Validaiton_1 no longer contains cell b4 but contains D1, A3:B3, A4, E2
		a1regions = a1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_1 after cut B4 to E2", 4, a1regions.size());
		Assert.assertFalse("Validation_1 not in B4 after cut B4 to E2", a1regions.contains(new CellRegion("b4"))); // b4
		Assert.assertTrue("Validation_1 in D1 after cut B4 to E2", a1regions.contains(new CellRegion("d1"))); // d1
		Assert.assertTrue("Validation_1 in E2 after cut B4 to E2", a1regions.contains(new CellRegion("e2"))); // e2
		Assert.assertTrue("Validation_1 in A4 after cut B4 to E2", a1regions.contains(new CellRegion("a4"))); // a4
		Assert.assertTrue("Validation_1 in A3:B3 after cut B4 to E2", a1regions.contains(new CellRegion("a3:b3"))); // a3:b3
		
		// 4.2. Validation_2 no longer contains cell E2 but contains D2 and E1 
		d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_2", 2, d1regions.size());
		Assert.assertFalse("Validation_2 in E2", d1regions.contains(new CellRegion("e2"))); // e2
		Assert.assertTrue("Validation_2 in D2", d1regions.contains(new CellRegion("d2"))); // d2
		Assert.assertTrue("Validation_2 in E1", d1regions.contains(new CellRegion("e1"))); // e1
		
		// 5. Copy E1 to Merged cell G1(G1:G4)
		Range e1 = Ranges.range(sheet, "e1");
		e1.paste(Ranges.range(sheet, "g1"));
		
		// Validation_2 still contains only D2 and E1 (copy and paste will not paste validation)
		d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_2", 2, d1regions.size());
		Assert.assertTrue("Validation_2 in D2", d1regions.contains(new CellRegion("d2"))); // d2
		Assert.assertTrue("Validation_2 in E1", d1regions.contains(new CellRegion("e1"))); // e1
		Assert.assertFalse("Validation_2 in G1", d1regions.contains(new CellRegion("g1"))); // g1
		
		// 6. Cut E1 to Merged cell G1(G1:G4)
		e1.paste(Ranges.range(sheet, "g1"), true);
		
		// 6.1. Validation_2 will be pasted to G1 which contains only D2 and G1; merged cell should be unmerged
		d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_2", 2, d1regions.size());
		Assert.assertTrue("Validation_2 in D2", d1regions.contains(new CellRegion("d2"))); // d2
		Assert.assertFalse("Validation_2 in E1", d1regions.contains(new CellRegion("e1"))); // e1
		Assert.assertTrue("Validation_2 in G1", d1regions.contains(new CellRegion("g1"))); // g1
	}

	@Test
	public void testZSS696CopyValidation() {
		Object[] books = _loadBooks(this, "book/696-cut-paste-mergecell.xlsx");
		Book book = (Book) books[0];
		Sheet sheet = book.getSheet("Sheet1");
		
		// 1. Validation_2 in D1:E2
		SSheet ssheet = sheet.getInternalSheet();
		SDataValidation d1dv = ssheet.getDataValidation(0, 3); // validation in  d1
		Set<CellRegion> d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions Validation_2", 1, d1regions.size());
		Assert.assertTrue("Validation_2 in D1:E2", d1regions.contains(new CellRegion("d1:e2"))); // d1:e2

		// 2. Copy E1 to Merged cell G1(G1:G4)
		Range e1 = Ranges.range(sheet, "e1");
		e1.paste(Ranges.range(sheet, "g1"));
		
		// 2.1 cannot copy paste Validation_2 to merged cell.
		d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_2", 1, d1regions.size());
		Assert.assertTrue("Validation_2 in D1:E2", d1regions.contains(new CellRegion("d1:e2"))); // d1:e2
		Assert.assertTrue("G1:G4 is a merged cell", Ranges.range(sheet, "g1:g4").isMergedCell());
		
		// 3. Cut E1 to Merged cell G1(G1:G4)
		e1.paste(Ranges.range(sheet, "g1"), true);
		
		// 3.1. Validation_2 will be pasted to G1; D1, D2:E2, G1
		d1regions = d1dv.getRegions();
		Assert.assertEquals("Number of regions in Validation_2", 3, d1regions.size());
		Assert.assertTrue("Validation_2 in D2", d1regions.contains(new CellRegion("d1"))); // d1
		Assert.assertFalse("Validation_2 in E1", d1regions.contains(new CellRegion("e1"))); // e1
		Assert.assertTrue("Validation_2 in E1", d1regions.contains(new CellRegion("d2:e2"))); // d2:e2
		Assert.assertTrue("Validation_2 in G1", d1regions.contains(new CellRegion("g1"))); // g1
		Assert.assertTrue("G1:G4 is an unmerged cell", !Ranges.range(sheet, "g1:g4").isMergedCell());
	}

	private static Object[] _loadBooks(Object base,String respath) {
		if(base==null){
			base = Util.class;
		}
		if(!(base instanceof Class)){
			base = base.getClass();
		}
		
		@SuppressWarnings("rawtypes")
		final InputStream is = ((Class)base).getResourceAsStream(respath);
		try {
			int index = respath.lastIndexOf("/");
			String bookName = index==-1?respath:respath.substring(index+1);
			XlsxImporter imp = new XlsxImporter();
			return new Object[] {new BookImpl(new SimpleRef<SBook>(imp.imports(is, bookName))), imp.getWorkbook()};
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(),e);
		} finally {
			if (is != null) {
				try {
					is.close();
				} catch (IOException e) {
				}
			}
		}
	}
}

@SuppressWarnings("serial")
final class XlsxExporter extends ExcelXlsxExporter {
	public Workbook getWorkbook() {
		return workbook;
	}
}

final class XlsxImporter extends ExcelXlsxImporter {
	public Workbook getWorkbook() {
		return workbook;
	}
}
