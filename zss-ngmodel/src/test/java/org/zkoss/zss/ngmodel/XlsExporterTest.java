package org.zkoss.zss.ngmodel;

import java.util.Iterator;

import junit.framework.Assert;

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
	public void exportWidthSplitTest() {
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<NColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		Assert.assertNull(sheet1.getColumnArray(0));
		
		sheet1.setupColumnArray(0, 10).setWidth(10);
		sheet1.setupColumnArray(11, 255);
		arrays = sheet1.getColumnArrayIterator();
		NColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(10, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(255, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
	}
	
	//unsupported features
	@Override @Ignore("not supported")
	public void areaChart() {
	}
	
	@Override @Ignore("not supported")
	public void barChart() {
	}
	
	@Override @Ignore("not supported")
	public void bubbleChart() {
	}
	
	@Override @Ignore("not supported")
	public void columnChart() {
	}
	
	@Override @Ignore("not supported")
	public void doughnutChart() {
	}
	
	@Override @Ignore("not supported")
	public void lineChart() {
	}
	
	@Override @Ignore("not supported")
	public void pieChart() {
	}

	@Override @Ignore("not supported")
	public void scatterChart() {
	}
	
	@Override @Ignore("not supported")
	public void picture() {
	}
	
	@Override @Ignore("not supported")
	public void validation() {
	}
}
