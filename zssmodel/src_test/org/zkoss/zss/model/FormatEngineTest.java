package org.zkoss.zss.model;

import java.awt.Color;
import java.util.Calendar;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.format.CellFormat;
import org.zkoss.poi.ss.format.CellFormatResult;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.impl.BookImpl;
import org.zkoss.zss.model.impl.ColorImpl;
import org.zkoss.zss.model.impl.sys.FormatEngineImpl;
import org.zkoss.zss.model.sys.format.FormatContext;
import org.zkoss.zss.model.sys.format.FormatEngine;
import org.zkoss.zss.model.sys.format.FormatResult;

public class FormatEngineTest {

	static private FormatContext formatContext;
	static private FormatEngine formatEngine;
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}	
	private SCell cell;
	
	@BeforeClass
	public static void init(){
		formatContext = new FormatContext(Locale.US);
		formatEngine = new FormatEngineImpl();
	}
	
	@Before
	public void beforeTest(){
		cell = createCell();
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	
	/* Number */
	
	@Test
	public void positiveNumber(){
		cell.setValue(new Integer(123456789));
		cell.getCellStyle().setDataFormat("0.00_);[red](0.00)");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("123456789.00 ", result.getText());
	}
	
	@Test
	public void negativeNumber(){
		cell.setValue(new Integer(-123456789));
		cell.getCellStyle().setDataFormat("0.00_);[red](0.00)");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("(123456789.00)", result.getText());
	}

	@Test
	public void thousandSeparator(){
		cell.setValue(new Integer(123456789));
		cell.getCellStyle().setDataFormat("#,##0.00_);(#,##0.00)");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("123,456,789.00 ", result.getText());
	
	}
	
	@Test
	public void scaleBy1000(){
		cell.setValue(new Integer(12000));
		cell.getCellStyle().setDataFormat("#.0,");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("12.0", result.getText());
	}

	
	/* Currency */
	@Test
	public void currency(){
		cell.setValue(new Double(-1234567890));
		cell.getCellStyle().setDataFormat("$#,##0.00_);($#,##0.00)");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("($1,234,567,890.00)", result.getText());
	}
	
	
	/* Accounting */
	@Test
	public void accounting(){
		cell.setValue(new Double(1234567890));
		cell.getCellStyle().setDataFormat("_-\"NT$\"* #,##0.00_ ;_-\"NT$\"* -#,##0.00 ;_-\"NT$\"* \"-\"??_ ;_-@_ ");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals(" NT$1,234,567,890.00 ", result.getText());
	}
	
	
	/* Date */
	@Test
	public void noLeadingZeroDate(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, 2013);
		calendar.set(Calendar.MONTH, 8);
		calendar.set(Calendar.DAY_OF_MONTH, 3);
		
		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("yyyy/m/d");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getDateValue(), formatContext);
		Assert.assertEquals("2013/9/3", result.getText());
	}
	
	@Test
	public void leadingZeroDate(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, 2013);
		calendar.set(Calendar.MONTH, 8);
		calendar.set(Calendar.DAY_OF_MONTH, 3);
		
		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("yyyy/mm/dd");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getDateValue(), formatContext);
		Assert.assertEquals("2013/09/03", result.getText());
	}
	
	@Test
	public void abbreviatedDate(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, 2013);
		calendar.set(Calendar.MONTH, 8);
		calendar.set(Calendar.DAY_OF_MONTH, 3);
		
		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("yy/mmm/ddd");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("13/Sep/Tue", result.getText());
	}	
		
	@Test
	public void fullNameDate(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, 2013);
		calendar.set(Calendar.MONTH, 8);
		calendar.set(Calendar.DAY_OF_MONTH, 3);
		
		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("yy/mmmm/dddd");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getDateValue(), formatContext);
		Assert.assertEquals("13/September/Tuesday", result.getText());
	}		
	
	@Test
	public void dateTime(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, 2013);
		calendar.set(Calendar.MONTH, 9);
		calendar.set(Calendar.DAY_OF_MONTH, 30);
		calendar.set(Calendar.HOUR_OF_DAY, 0);
		calendar.set(Calendar.MINUTE, 0);

		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("yyyy/m/d h:mm;@");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getDateValue(), formatContext);
		Assert.assertEquals("2013/10/30 0:00", result.getText());
	}

	
	@Test
	public void TaiwanDate(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.YEAR, 2013);
		calendar.set(Calendar.MONTH, 9);
		calendar.set(Calendar.DAY_OF_MONTH, 30);
		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("yyyy/mmm/ddd");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getDateValue(), new FormatContext(Locale.TAIWAN));
		Assert.assertEquals("2013/十月/星期三", result.getText());
	}
	
	
	/* Time */
	@Test
	public void time(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.HOUR_OF_DAY, 11);
		calendar.set(Calendar.MINUTE, 10);
		calendar.set(Calendar.SECOND, 50);

		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("h:mm:ss AM/PM;@");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getDateValue(), formatContext);
		Assert.assertEquals("11:10:50 AM", result.getText());
	}

	@Test
	public void unsupportedAbbriviated12HourTime(){
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.HOUR_OF_DAY, 11);
		calendar.set(Calendar.MINUTE, 10);
		calendar.set(Calendar.SECOND, 50);

		cell.setValue(calendar.getTime());
		cell.getCellStyle().setDataFormat("a/p");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("a/p", result.getText());
	}
	
	@Test
	public void elapsedTime(){
		cell.setValue(new Double(2.46585648148148));
		cell.getCellStyle().setDataFormat("[h]");
		FormatResult result1 = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("59", result1.getText());
		
		cell.getCellStyle().setDataFormat("[m]");
		FormatResult result2 = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("3550", result2.getText());
		
		cell.getCellStyle().setDataFormat("[s]");
		FormatResult result3 = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("213049", result3.getText());
	}
	
	
	/* Percentage*/
	@Test
	public void percentage(){
		cell.setValue(new Double(0.98585));
		cell.getCellStyle().setDataFormat("0.00%");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("98.59%", result.getText());
	}
	
	/* Fraction */
	@Test
	public void fraction(){
		cell.setValue(new Double(0.330858961));
		cell.getCellStyle().setDataFormat("# ???/???");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals(" 312/943", result.getText());
	}
	
	@Test
	public void fractionHundredths(){
		cell.setValue(new Double(0.3));
		cell.getCellStyle().setDataFormat("# ??/100");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals(" 30/100", result.getText());
	}
	
	
	/* Scientific */
	@Test
	public void scientific(){
		cell.setValue(new Double(123456789));
		cell.getCellStyle().setDataFormat("0.00E+00");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("1.23E+08", result.getText());
	}
	
	
	/* Text */
	@Test
	public void text(){
		cell.setValue(new Double(123456789));
		cell.getCellStyle().setDataFormat("@");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("123456789", result.getText());
	}
	
	
	/* Special */
	@Test
	public void zipCode(){
		cell.setValue(new Double(156769884));
		cell.getCellStyle().setDataFormat("0000000-0");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("15676988-4", result.getText());
	}
	
	@Test
	public void phone(){
		cell.setValue(new Double(21234567));
		cell.getCellStyle().setDataFormat("[<=9999999]###-####;(0#) ###-####");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("(02) 123-4567", result.getText());
	}
	
	
	/* Custom - color */
	@Test
	public void color(){
		cell.setValue(new Integer(12000));
		cell.getCellStyle().setDataFormat("[red]#.0,");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals(ColorImpl.RED, result.getColor());
	}
	

	
	/* General */
	@Test
	public void generalFormat(){
		cell.setValue(new Double(-1234567890));
		cell.getCellStyle().setDataFormat("General");
		FormatResult result = formatEngine.format(cell.getCellStyle().getDataFormat(),cell.getValue(), formatContext);
		Assert.assertEquals("-1234567890", result.getText());
	}
	
	/* utility methods */
	private SCell createCell(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = book.createSheet("Sheet1");
		return sheet1.getCell(0,0);
	}
}
