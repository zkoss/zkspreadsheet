package org.zkoss.zss.api.impl;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.Sheet;

/**
 * For Input of date,it's always guess the input text by current local, for example. if locale is TW, then guess from format 'yyyy/m/d', if local is US, then 
 * guess from format 'm/d/yyyy'. 
 * If you input a wrong format text, the cell value will become a String and format becomes 'General'
 * 
 * For Displaying a date, it depends on cell's data format. 
 * However there is a special format 'm/d/yyyy', when you set it to a cell, the display will also depends on current locale. 
 * for example, if locale is TW, the display will be formated to 'yyyy/m/d' and 
 * if local is US, the display will be formated to 'm/d/yyyy'  
 * 
 * When setting the data format, it also guesses the input format which depends on current locale and transfer it to special format('m/d/yyyy')
 * for example, if locale is TW, setting format 'yyyy/m/d' will be transfer to special format 'm/d/yyyy' 
 *  
 * 
 * 
 * @author dennis
 *
 */
public class DateInputTest {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}
	
	@Test
	public void testDateDisplaySpecia(){
		testDateDisplaySpecial1(Util.loadBook(this, "book/blank.xlsx"));
		testDateDisplaySpecial1(Util.loadBook(this, "book/blank.xls"));
		testDateDisplaySpecial2(Util.loadBook(this, "book/blank.xlsx"));
		testDateDisplaySpecial2(Util.loadBook(this, "book/blank.xls"));
		testDateDisplaySpecial3(Util.loadBook(this, "book/blank.xlsx"));
		testDateDisplaySpecial3(Util.loadBook(this, "book/blank.xls"));
	}
	public void testDateDisplaySpecial1(Book book){
			
			
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		r = Ranges.range(sheet,"A1");
		r.setCellEditText("2013/10/25"); // data format in TW
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());
		
		// special format the will cause get format text depends on current locale
		CellOperationUtil.applyDataFormat(r, "m/d/yyyy");
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat());
		Assert.assertEquals("2013/10/25", r.getCellFormatText());// display depends on locale
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());// edit always depends on local
		//US
		Setup.setZssContentLocale(Locale.US);
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat());
		//*** display depends on locale ***//
		Assert.assertEquals("10/25/2013", r.getCellFormatText());// display depends on locale
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());// edit always depends on local
		//UK
		Setup.setZssContentLocale(Locale.UK);
		Assert.assertEquals("dd-mm-yyyy", r.getCellStyle().getDataFormat());
		//*** display depends on locale ***//
		Assert.assertEquals("25-10-2013", r.getCellFormatText());// display depends on locale
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("25-10-2013", r.getCellEditText());// edit always depends on local		
	}
	
	public void testDateDisplaySpecial2(Book book){
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		r = Ranges.range(sheet,"A1");
		r.setCellEditText("2013/10/25"); // data format in TW
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());
		
		// special format the will cause get format text depends on current locale
		CellOperationUtil.applyDataFormat(r, "yyyy/m/d"); // when transfer to special m/d/yyyy
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat());
		Assert.assertEquals("2013/10/25", r.getCellFormatText());// display depends on locale
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());// edit always depends on local
		//US
		Setup.setZssContentLocale(Locale.US);
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat());
		//*** display depends on locale ***//
		Assert.assertEquals("10/25/2013", r.getCellFormatText());// display depends on locale
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());// edit always depends on local
		//UK
		Setup.setZssContentLocale(Locale.UK);
		Assert.assertEquals("dd-mm-yyyy", r.getCellStyle().getDataFormat());
		//*** display depends on locale ***//
		Assert.assertEquals("25-10-2013", r.getCellFormatText());// display depends on locale
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("25-10-2013", r.getCellEditText());// edit always depends on local		
	}
	
	public void testDateDisplaySpecial3(Book book){
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssContentLocale(Locale.US);
		r = Ranges.range(sheet,"A1");
		r.setCellEditText("10/25/2013"); // data format in TW
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("10/25/2013", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());
		
		// special format the will cause get format text depends on current locale
		CellOperationUtil.applyDataFormat(r, "yyyy/m/d"); // will not transfer to special format
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat());
		Assert.assertEquals("2013/10/25", r.getCellFormatText());// display depends on format only
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());// edit always depends on local
		//US
		Setup.setZssContentLocale(Locale.US);
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat());
		Assert.assertEquals("2013/10/25", r.getCellFormatText());// display depends on format only
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());// edit always depends on local
		//UK
		Setup.setZssContentLocale(Locale.UK);
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat());
		Assert.assertEquals("2013/10/25", r.getCellFormatText());// display depends on format only
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("25-10-2013", r.getCellEditText());// edit always depends on local		
	}
	
	@Test
	public void testDateDisplayNormal(){
		testDateDisplayNormal1(Util.loadBook(this, "book/blank.xlsx"));
		testDateDisplayNormal1(Util.loadBook(this, "book/blank.xls"));
		testDateDisplayNormal2(Util.loadBook(this, "book/blank.xlsx"));
		testDateDisplayNormal2(Util.loadBook(this, "book/blank.xls"));
	
	}
	public void testDateDisplayNormal1(Book book){	
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		r = Ranges.range(sheet,"A1");
		r.setCellEditText("2013/10/25"); // data format in TW
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText()); // edit always depends on local
		
		// normal format
		CellOperationUtil.applyDataFormat(r, "m/d");
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		Assert.assertEquals("m/d", r.getCellStyle().getDataFormat());
		Assert.assertEquals("10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());// edit always depends on local
		
		//US
		Setup.setZssContentLocale(Locale.US);
		Assert.assertEquals("m/d", r.getCellStyle().getDataFormat());
		//*** display doesn't depends on locale ***//
		Assert.assertEquals("10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());// edit always depends on local

		//UK
		Setup.setZssContentLocale(Locale.UK);
		Assert.assertEquals("m/d", r.getCellStyle().getDataFormat());
		//*** display doesn't depends on locale ***//
		Assert.assertEquals("10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("25-10-2013", r.getCellEditText());// edit always depends on local		
	}
	
	public void testDateDisplayNormal2(Book book){
		
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		r = Ranges.range(sheet,"A1");
		r.setCellEditText("2013/10/25"); // data format in TW
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText()); // edit always depends on local
		
		// normal format
		CellOperationUtil.applyDataFormat(r, "d/m");
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		Assert.assertEquals("d/m", r.getCellStyle().getDataFormat());
		Assert.assertEquals("25/10", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());// edit always depends on local
		
		//US
		Setup.setZssContentLocale(Locale.US);
		Assert.assertEquals("d/m", r.getCellStyle().getDataFormat());
		//*** display doesn't depends on locale ***//
		Assert.assertEquals("25/10", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());// edit always depends on local

		//UK
		Setup.setZssContentLocale(Locale.UK);
		Assert.assertEquals("d/m", r.getCellStyle().getDataFormat());
		//*** display doesn't depends on locale ***//
		Assert.assertEquals("25/10", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("25-10-2013", r.getCellEditText());// edit always depends on local		
	}
	
	@Test
	public void testDateInput(){
		testDateInput1(Util.loadBook(this, "book/blank.xlsx"));
		testDateInput1(Util.loadBook(this, "book/blank.xls"));
	}
	public void testDateInput1(Book book){
		Sheet sheet = book.getSheetAt(0);
		Range r;
		
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		r = Ranges.range(sheet,"A1");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		r.setCellEditText("2013/10/25"); // data format in TW
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("2013/10/25", r.getCellEditText());
		
		r = Ranges.range(sheet,"A2");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		r.setCellEditText("10/25/2013"); // data format in US
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());//format doesn't chenage
		Assert.assertEquals("10/25/2013", r.getCellFormatText());
		Assert.assertEquals(CellType.STRING, r.getCellData().getType());//it is string
		Assert.assertEquals("10/25/2013", r.getCellEditText());
		
		
		//US
		Setup.setZssContentLocale(Locale.US);
		
		r = Ranges.range(sheet,"B1");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		r.setCellEditText("10/25/2013"); // data format in US
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("10/25/2013", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("10/25/2013", r.getCellEditText());
		
		r = Ranges.range(sheet,"B2");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		r.setCellEditText("2013/10/25"); // data format in TW , not current local
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());//format doesn't change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.STRING, r.getCellData().getType());//it is string
		Assert.assertEquals("2013/10/25", r.getCellEditText());

		//UK
		Setup.setZssContentLocale(Locale.UK);
		
		r = Ranges.range(sheet,"C1");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		r.setCellEditText("25-10-2013"); // data format in US
		Assert.assertEquals("dd-mm-yyyy", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals("25-10-2013", r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals("25-10-2013", r.getCellEditText());
		
		r = Ranges.range(sheet,"C2");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		r.setCellEditText("2013/10/25"); // data format in TW , not current local
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());//format doesn't change
		Assert.assertEquals("2013/10/25", r.getCellFormatText());
		Assert.assertEquals(CellType.STRING, r.getCellData().getType());//it is string
		Assert.assertEquals("2013/10/25", r.getCellEditText());		
	}

	
	
	
	@Test
	//ZSS-495
	public void testInputDateObject(){
		testInputDateObjectTW(Util.loadBook(Issue400Test.class, "book/blank.xlsx"));
		testInputDateObjectUS(Util.loadBook(Issue400Test.class, "book/blank.xlsx"));
	}
	
	public void testInputDateObjectTW(Book book){	
		Sheet sheet = book.getSheetAt(0);
		Range r;
		Date today = new Date();
		
		String todayStr;
		String todayStrLong;
		//TW
		Setup.setZssContentLocale(Locale.TAIWAN);
		r = Ranges.range(sheet,"A1");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		CellOperationUtil.applyDataFormat(r, "m/d/yyyy"); 
		todayStr = new SimpleDateFormat("yyyy/MM/dd",Setup.getZssContextLocale()).format(today);
		todayStrLong = new SimpleDateFormat("yyyy/MM/dd a hh:mm:ss",Setup.getZssContextLocale()).format(today); //it is hh:mm:ss in TW
		
		r.setCellEditText(todayStr);
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals(todayStr, r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals(todayStr, r.getCellEditText());
		
		
		r.getCellData().setValue(today);
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); 
		Assert.assertEquals(todayStr, r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals(todayStrLong, r.getCellEditText());
		
		r.getCellData().setValue(Util.getDateOnly(today));
		Assert.assertEquals("yyyy/m/d", r.getCellStyle().getDataFormat()); 
		Assert.assertEquals(todayStr, r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals(todayStr, r.getCellEditText());
	}
	
	public void testInputDateObjectUS(Book book){	
		Sheet sheet = book.getSheetAt(0);
		Range r;
		Date today = new Date();
		String todayStr,todayStrLong;
		//US
		Setup.setZssContentLocale(Locale.US);
		r = Ranges.range(sheet,"A1");
		Assert.assertEquals("General", r.getCellStyle().getDataFormat());
		CellOperationUtil.applyDataFormat(r, "m/d/yyyy"); 
		todayStr = Util.getDateFormat("MM/dd/yyyy",Setup.getZssContextLocale()).format(today);
		todayStrLong = Util.getDateFormat("MM/dd/yyyy h:mm:ss a",Setup.getZssContextLocale()).format(today); // it is h:mm:ss in US
		
		r.setCellEditText(todayStr);
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat()); //format change
		Assert.assertEquals(todayStr, r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals(todayStr, r.getCellEditText());
		
		
		r.getCellData().setValue(today);
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat()); 
		Assert.assertEquals(todayStr, r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals(todayStrLong, r.getCellEditText());
		
		r.getCellData().setValue(Util.getDateOnly(today));
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat()); 
		Assert.assertEquals(todayStr, r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals(todayStr, r.getCellEditText());
		
		//test again
		r.getCellData().setValue(Util.getDateOnly(today));
		Assert.assertEquals("m/d/yyyy", r.getCellStyle().getDataFormat()); 
		Assert.assertEquals(todayStr, r.getCellFormatText());
		Assert.assertEquals(CellType.NUMERIC, r.getCellData().getType());
		Assert.assertEquals(todayStr, r.getCellEditText());
	}
}
