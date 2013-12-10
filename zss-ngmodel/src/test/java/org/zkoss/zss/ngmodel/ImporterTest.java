package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;

import org.junit.Test;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.ExcelImportFactory;

public class ImporterTest {
	
	
	@Test
	public void loadXlsxBook() {
		final InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		NImporter importer = new ExcelImportFactory().createImporter();
		NBook book = null;
		try {
			book = importer.imports(is, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		assertEquals(book.getBookName(), "XSSFBook");

		// 3 sheet
		assertEquals(book.getNumOfSheet(), 3);

		NSheet sheet1 = book.getSheet(0);
		assertEquals(sheet1.getSheetName(), "First");
		NSheet sheet2 = book.getSheet(1);
		assertEquals(sheet2.getSheetName(), "Second");
		NSheet sheet3 = book.getSheet(2);
		assertEquals(sheet3.getSheetName(), "Third");
	}
	
	@Test
	public void cellsTest() {
		final InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		NImporter importer = new ExcelImportFactory().createImporter();
		NBook book = null;
		try {
			book = importer.imports(is, "XSSFBook");
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
		assertEquals(0.61, row2.getCell(2).getNumberValue().doubleValue(), 0.01);
		
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
	}

}
