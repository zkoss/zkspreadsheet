package org.zkoss.zss.ngmodel;

import static org.junit.Assert.fail;

import java.io.File;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.ngapi.impl.ExcelExportFactory;
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
	public void book(){
		try{
			xlsxExporter.export(bookUnderTest, new File("book/"+exportFileName));
		}catch (Exception e) {
			e.printStackTrace();
			fail(e.toString());
		}
	}
}

