package org.zkoss.zss.ngmodel;

import java.awt.Color;
import java.util.Calendar;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.ngmodel.impl.BookImpl;
import org.zkoss.zss.ngmodel.impl.sys.FormatEngineImpl;
import org.zkoss.zss.ngmodel.sys.format.FormatContext;
import org.zkoss.zss.ngmodel.sys.format.FormatEngine;
import org.zkoss.zss.ngmodel.sys.format.FormatResult;

public class CellFormatTest {

	static private FormatContext formatContext;
	static private FormatEngine formatEngine;
	
	@BeforeClass
	public static void init(){
		formatContext = new FormatContext(Locale.US);
		formatEngine = new FormatEngineImpl();
	}
	
	@Test
	public void numberTest(){
		NBook book = new BookImpl("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		sheet1.getCell(0,0).setValue(new Integer(12000));
		sheet1.getCell(0,0).getCellStyle().setDataFormat("#.0,");
		NCell cell = sheet1.getCell(0,0);
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("12.0", result.getText());
	}

	@Test
	public void currencyTest(){
		NBook book = new BookImpl("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		sheet1.getCell(0,0).setValue(new Double(-1234567890));
		sheet1.getCell(0,0).getCellStyle().setDataFormat("$#,##0.00_);($#,##0.00)");
		NCell cell = sheet1.getCell(0,0);
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("($1,234,567,890.00)", result.getText());
	}
	
	@Test
	public void dateTest(){
		NBook book = new BookImpl("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, 2013);
		calendar.set(Calendar.MONTH, 9);
		calendar.set(Calendar.DAY_OF_MONTH, 30);
		sheet1.getCell(0,0).setValue(calendar.getTime());
		sheet1.getCell(0,0).getCellStyle().setDataFormat("yyyy/m/d");
		NCell cell = sheet1.getCell(0,0);
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("2013/10/30", result.getText());
	}
//	
//	@Test
//	public void dateTwTest(){
//		Calendar calendar = Calendar.getInstance();
//		calendar.set(Calendar.YEAR, 2013);
//		calendar.set(Calendar.MONTH, 9);
//		calendar.set(Calendar.DAY_OF_MONTH, 30);
//		CellFormatResult result = CellFormat.getInstance("yyyy/m/d", Locale.TAIWAN).apply(calendar.getTime());
//		assertEquals("2013/10/30", result.text);
//	}
//	
//	@Test
//	public void zipCodeTest(){
//		CellFormatResult result = CellFormat.getInstance("0000000-0", Locale.US).apply(new Double(156769884));
//		assertEquals("15676988-4", result.text);
//	}
	
	@Test
	public void colorTest(){
		NBook book = new BookImpl("book1");
		NSheet sheet1 = book.createSheet("Sheet1");
		sheet1.getCell(0,0).setValue(new Integer(12000));
		sheet1.getCell(0,0).getCellStyle().setDataFormat("[red]#.0,");
		NCell cell = sheet1.getCell(0,0);
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals(Color.RED.toString(), result.getColor());
	}
	
//	//"general" format
//	@Test
//	public void generalFormatTest(){
//		CellFormatResult result = CellFormat.getInstance("General", Locale.US).apply(new Double(-1234567890));
//		assertEquals("-1234567890)", result.text);
//	}
}
