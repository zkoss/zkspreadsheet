package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.MalformedURLException;
import java.net.URL;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.ExcelExportFactory;
import org.zkoss.zss.ngapi.impl.ExcelImportFactory;
import org.zkoss.zss.ngapi.impl.NExcelXlsxExporter;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class ExporterTest {

	static private NBook bookUnderTest;
	static private String exportFileName = "exported.xlsx";
	static private NExcelXlsxExporter xlsxExporter = (NExcelXlsxExporter) new ExcelExportFactory(ExcelExportFactory.Type.XLSX).createExporter();

	@BeforeClass
	static public void createBookForExport() {
		String bookName = "book for export";
		bookUnderTest = NBooks.createBook(bookName);
	}

	@Test
	public void exportBorderFileTest() throws IOException {
		// Import at first
		InputStream is = ImporterTest.class.getResourceAsStream("book/cell_borders.xlsx");
		NImporter importer = new ExcelImportFactory().createImporter();
		NBook book = null;
		try {
			book = importer.imports(is, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			is.close();
		}

		File outFile = null;
		try {
			outFile = new File(getClass().getProtectionDomain().getCodeSource().getLocation().getPath() + "/org/zkoss/zss/ngmodel/book/" + exportFileName);
			outFile.createNewFile();
			xlsxExporter.export(book, outFile);
		} catch (Exception e) {
			e.printStackTrace();
			fail(e.toString());
		}
	}

	@Test
	public void exportBorderModelTest() {
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		NCell cell1 = sheet1.getCell(1, 1);
		NCell cell2 = sheet1.getCell(1, 2);
		NCell cell3 = sheet1.getCell(1, 3);

		cell1.setStringValue("hair");
		cell2.setStringValue("dot");
		cell3.setStringValue("dash");

		NCellStyle style1 = book.createCellStyle(true);
		style1.setBorderBottom(BorderType.HAIR);
		cell1.setCellStyle(style1);

		NCellStyle style2 = book.createCellStyle(true);
		style2.setBorderBottom(BorderType.DOTTED);
		cell2.setCellStyle(style2);

		NCellStyle style3 = book.createCellStyle(true);
		style3.setBorderBottom(BorderType.DASHED);
		cell3.setCellStyle(style3);

		NCell cell21 = sheet1.getCell(2, 1);
		NCell cell22 = sheet1.getCell(2, 2);
		NCell cell23 = sheet1.getCell(2, 3);

		NCellStyle style21 = book.createCellStyle(true);
		style21.setBorderTop(BorderType.NONE);
		cell21.setCellStyle(style21);
		NCellStyle style22 = book.createCellStyle(true);
		style22.setBorderTop(BorderType.NONE);
		cell22.setCellStyle(style22);
		NCellStyle style23 = book.createCellStyle(true);
		style23.setBorderTop(BorderType.NONE);
		cell23.setCellStyle(style23);

		File outFile = null;
		try {
			outFile = new File(getClass().getProtectionDomain().getCodeSource().getLocation().getPath() + "/org/zkoss/zss/ngmodel/book/" + exportFileName);
			outFile.createNewFile();
			xlsxExporter.export(book, outFile);
		} catch (Exception e) {
			e.printStackTrace();
			fail(e.toString());
		}
	}

	@Test
	public void book() throws IOException {

		// Import at first
		InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		NImporter importer = new ExcelImportFactory().createImporter();
		NBook book = null;
		try {
			book = importer.imports(is, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			is.close();
		}

		File outFile = null;
		try {
			outFile = new File(getClass().getProtectionDomain().getCodeSource().getLocation().getPath() + "/org/zkoss/zss/ngmodel/book/" + exportFileName);
			outFile.createNewFile();
			xlsxExporter.export(book, outFile);
		} catch (Exception e) {
			e.printStackTrace();
			fail(e.toString());
		}

		// is = new FileInputStream(outFile);
		// importer = new ExcelImportFactory().createImporter();
		// book = null;
		// try {
		// book = importer.imports(is, "XSSFBook");
		// } catch (IOException e) {
		// e.printStackTrace();
		// }finally {
		// is.close();
		// }
		//
		// NSheet sheet = book.getSheetByName("cell-border");
		// assertEquals(BorderType.DOTTED, sheet.getCell(4,
		// 2).getCellStyle().getBorderBottom());
		// assertEquals(BorderType.DASHED, sheet.getCell(4,
		// 3).getCellStyle().getBorderBottom());
	}
}
