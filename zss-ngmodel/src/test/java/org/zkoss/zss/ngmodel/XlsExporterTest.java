package org.zkoss.zss.ngmodel;

import java.io.File;

import org.junit.*;
import org.zkoss.zss.ngapi.impl.imexp.ExcelExportFactory;

public class XlsExporterTest extends ExporterTest {


	@Before
	public void setupTestFile(){
		IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/import.xls");
		CHART_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/chart.xls");
		PICTURE_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/picture.xls");
		EXPORTER_TYPE = ExcelExportFactory.Type.XLS;
	}
	
	@Override
	public void sheetTest() {
		NBook book4Test = ImExpTestUtil.loadBook(IMPORT_FILE_UNDER_TEST, "XSSFBook");
		File outFile = ImExpTestUtil.write(book4Test, EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
		sheetTest(book);
	}
	
	@Test
	public void definedName(){
		NBook book4Test = ImExpTestUtil.loadBook(ImporterTest.class.getResource("book/range.xls"), "XSSFBook");
		File outFile = ImExpTestUtil.write(book4Test, EXPORTER_TYPE);
		NBook book = ImExpTestUtil.loadBook(outFile, DEFAULT_BOOK_NAME);
	}
	
	//unsupported features
	@Override @Ignore
	public void areaChart() {
	}
	
	@Override @Ignore
	public void barChart() {
	}
	
	@Override @Ignore
	public void bubbleChart() {
	}
	
	@Override @Ignore
	public void columnChart() {
	}
	
	@Override @Ignore
	public void doughnutChart() {
	}
	
	@Override @Ignore
	public void lineChart() {
	}
	
	@Override @Ignore
	public void pieChart() {
	}

	@Override @Ignore
	public void scatterChart() {
	}
	
	@Override @Ignore
	public void picture() {
	}
	
	@Override @Ignore
	public void validation() {
	}
}
