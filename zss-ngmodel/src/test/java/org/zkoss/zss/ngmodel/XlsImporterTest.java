package org.zkoss.zss.ngmodel;

import org.junit.*;

public class XlsImporterTest extends ImporterTest {

	@BeforeClass
	static public void setupTestFile(){
		IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/import.xls");
		CHART_IMPORT_FILE_UNDER_TEST = ImporterTest.class.getResource("book/chart.xls");
	}
}
