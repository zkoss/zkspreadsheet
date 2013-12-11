package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.text.DateFormat;
import java.util.Locale;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.ExcelImportFactory;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;

public class ImporterTest {
	
	static private File fileUnderTest;
	private NImporter importer; 
	
	/**
	 * For exporter test to specify its exported file to test.
	 * @param file
	 */
	static public void setFileUnderTest(File file){
		fileUnderTest = file;
	}
	
	@BeforeClass
	static public void initialize(){
		try{
			fileUnderTest = new File(ImporterTest.class.getResource("book/import.xlsx").toURI());
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@Before
	public void prepare(){
		importer= new ExcelImportFactory().createImporter();
	}
	
	//API
	
	@Test
	public void importByInputStream(){
		InputStream streamUnderTest = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		NBook book = null;
		try {
			book = importer.imports(streamUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals(book.getBookName(), "XSSFBook");
		assertEquals(book.getNumOfSheet(), 3);
	}
	
	@Test
	public void importByUrl(){
		URL surlUnderTest = ImporterTest.class.getResource("book/import.xlsx");
		NBook book = null;
		try {
			book = importer.imports(surlUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals(book.getBookName(), "XSSFBook");
		assertEquals(book.getNumOfSheet(), 3);
	}
	
	@Test
	public void importByFile() {
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals(book.getBookName(), "XSSFBook");

		// 3 sheet
		assertEquals(book.getNumOfSheet(), 3);

		NSheet sheet1 = book.getSheet(0);
		assertEquals("Value", sheet1.getSheetName());
		NSheet sheet2 = book.getSheet(1);
		assertEquals("Style", sheet2.getSheetName());
		NSheet sheet3 = book.getSheet(2);
		assertEquals("Third", sheet3.getSheetName());
	}

	//content
	@Test
	public void cellValueTest() {
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		NSheet sheet = book.getSheet(0);
		//text
		NRow row = sheet.getRow(0);
		assertEquals(NCell.CellType.STRING, row.getCell(1).getType());
		assertEquals("B1", row.getCell(1).getStringValue());
		assertEquals("C1", row.getCell(2).getStringValue());
		assertEquals("D1", row.getCell(3).getStringValue());
		
		//number
		NRow row1 = sheet.getRow(1);
		assertEquals(NCell.CellType.NUMBER, row1.getCell(1).getType());
		assertEquals(123, row1.getCell(1).getNumberValue().intValue());
		assertEquals(123.45, row1.getCell(2).getNumberValue().doubleValue(), 0.01);
		
		//date
		NRow row2 = sheet.getRow(2);
		assertEquals(NCell.CellType.NUMBER, row2.getCell(1).getType());
		assertEquals(41618, row2.getCell(1).getNumberValue().intValue());
		assertEquals("Dec 10, 2013", DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.US).format(row2.getCell(1).getDateValue()));
		assertEquals(0.61, row2.getCell(2).getNumberValue().doubleValue(), 0.01);
		assertEquals("2:44:10 PM", DateFormat.getTimeInstance (DateFormat.MEDIUM, Locale.US).format(row2.getCell(2).getDateValue()));
		
		//formula
		NRow row3 = sheet.getRow(3);
		assertEquals(NCell.CellType.FORMULA, row3.getCell(1).getType());
		assertEquals("SUM(10,20)", row3.getCell(1).getFormulaValue());
		assertEquals("ISBLANK(B1)", row3.getCell(2).getFormulaValue());
		assertEquals("B1", row3.getCell(3).getFormulaValue());
		
		//error
		NRow row4 = sheet.getRow(4);
		assertEquals(NCell.CellType.ERROR, row4.getCell(1).getType());
		assertEquals(ErrorValue.INVALID_NAME, row4.getCell(1).getErrorValue().getCode());
		assertEquals(ErrorValue.INVALID_VALUE, row4.getCell(2).getErrorValue().getCode());
		
		//blank
		NRow row5 = sheet.getRow(5);
		assertEquals(NCell.CellType.BLANK, row5.getCell(1).getType());
		assertEquals("", row5.getCell(1).getStringValue());
	}

	@Test
	public void cellFontNameTest(){
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		NSheet sheet = book.getSheetByName("Style");
		assertEquals("Arial", sheet.getCell(3, 0).getCellStyle().getFont().getName());
		assertEquals("Arial Black", sheet.getCell(3, 1).getCellStyle().getFont().getName());
		assertEquals("Calibri", sheet.getCell(3, 2).getCellStyle().getFont().getName());
	}
	
	@Test
	public void cellFontStyleTest(){
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(NFont.Boldweight.BOLD, sheet.getCell(9, 0).getCellStyle().getFont().getBoldweight());
		assertTrue(sheet.getCell(9, 1).getCellStyle().getFont().isItalic());
		assertTrue(sheet.getCell(9, 2).getCellStyle().getFont().isStrikeout());
		assertEquals(NFont.Underline.SINGLE, sheet.getCell(9, 3).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.DOUBLE, sheet.getCell(9, 4).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.SINGLE_ACCOUNTING, sheet.getCell(9, 5).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.DOUBLE_ACCOUNTING, sheet.getCell(9, 6).getCellStyle().getFont().getUnderline());
		assertEquals(NFont.Underline.NONE, sheet.getCell(9, 7).getCellStyle().getFont().getUnderline());
		
		//height
		assertEquals(8, sheet.getCell(6, 0).getCellStyle().getFont().getHeightPoints());
		assertEquals(72, sheet.getCell(6, 3).getCellStyle().getFont().getHeightPoints());
		
		//type offset
		assertEquals(TypeOffset.SUPER, sheet.getCell(32, 1).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.SUB, sheet.getCell(32, 2).getCellStyle().getFont().getTypeOffset());
		assertEquals(TypeOffset.NONE, sheet.getCell(32, 3).getCellStyle().getFont().getTypeOffset());
	}
	
	@Test
	public void rowTest(){
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		NSheet sheet = book.getSheetByName("Style");
		assertEquals(28, sheet.getRow(0).getHeight());
		assertEquals(20, sheet.getRow(1).getHeight());
	}
	
	@Test
	public void columnTest(){
		NBook book = null;
		try {
			book = importer.imports(fileUnderTest, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		NSheet sheet = book.getSheetByName("Style");
		
		assertEquals(28, sheet.getColumn(0).getWidth());
		assertEquals(20, sheet.getColumn(1).getWidth());
	}
}
