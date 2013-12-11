package org.zkoss.zss.ngmodel;

import static org.junit.Assert.fail;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.ExcelExportFactory;
import org.zkoss.zss.ngapi.impl.ExcelImportFactory;
import org.zkoss.zss.ngapi.impl.NExcelXlsxExporter;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class ExporterTest {

	static private NBook bookUnderTest;
	static private String exportFileName = "exported.xlsx";
	static private NExcelXlsxExporter xlsxExporter = (NExcelXlsxExporter)new ExcelExportFactory(ExcelExportFactory.Type.XLSX).createExporter();
	
	@BeforeClass
	static public void createBookForExport(){
		String bookName = "book for export";
		bookUnderTest = new BookImpl(bookName); 
	}
	
	@Test
	public void book() throws MalformedURLException {

		// Import at first
		final InputStream is = ImporterTest.class.getResourceAsStream("book/import.xlsx");
		NImporter importer = new ExcelImportFactory().createImporter();
		NBook book = null;
		try {
			book = importer.imports(is, "XSSFBook");
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		try{
			File outFile = new File(getClass().getProtectionDomain().getCodeSource().getLocation().getPath() + "/org/zkoss/zss/ngmodel/book/" + exportFileName);
			outFile.createNewFile();
			xlsxExporter.export(book, outFile);
		}catch (Exception e) {
			e.printStackTrace();
			fail(e.toString());
		}
	}
}

