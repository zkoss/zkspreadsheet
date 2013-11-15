package org.zkoss.zss.ngmodel;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class ModelTest {

	@Test
	public void testException(){
		NBook book = new BookImpl();
		NSheet sheet1 = book.createSheet("Sheet1");
		Assert.assertEquals(1, book.getNumOfSheet());
		NSheet sheet2 = book.createSheet("Sheet2");
		Assert.assertEquals(2, book.getNumOfSheet());
		
		try{
			NSheet sheet = book.createSheet("Sheet2");
			Assert.fail("should get exception");
		}catch(InvalidateModelOpException x){}
		
		Assert.assertEquals(2, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(0));
		Assert.assertEquals(sheet2, book.getSheetAt(1));
		Assert.assertEquals(sheet1, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(null, book.getSheetByName("Sheet3"));
		
		book.deleteSheet(sheet1);
		
		Assert.assertEquals(1, book.getNumOfSheet());
		Assert.assertEquals(sheet2, book.getSheetAt(0));
		Assert.assertEquals(null, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(null, book.getSheetByName("Sheet3"));
		
		try{
			book.deleteSheet(sheet1);
			Assert.fail("should get exception");
		}catch(InvalidateModelOpException x){}//ownership
		
		try{
			book.createSheet("Sheet3", sheet1);
			Assert.fail("should get exception");
		}catch(InvalidateModelOpException x){}//ownership
		
		try{
			book.moveSheetTo(sheet1, 0);
			Assert.fail("should get exception");
		}catch(InvalidateModelOpException x){}//ownership
		
		sheet1 = book.createSheet("Sheet1");
		
		Assert.assertEquals(2, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(1));
		Assert.assertEquals(sheet2, book.getSheetAt(0));
		Assert.assertEquals(sheet1, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(null, book.getSheetByName("Sheet3"));
		
		NSheet sheet3 = book.createSheet("Sheet3");
		
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(1));
		Assert.assertEquals(sheet2, book.getSheetAt(0));
		Assert.assertEquals(sheet3, book.getSheetAt(2));
		Assert.assertEquals(sheet1, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(sheet3, book.getSheetByName("Sheet3"));
		
		
		book.moveSheetTo(sheet1, 0);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(0));
		Assert.assertEquals(sheet2, book.getSheetAt(1));
		Assert.assertEquals(sheet3, book.getSheetAt(2));
		
		book.moveSheetTo(sheet1, 1);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(1));
		Assert.assertEquals(sheet2, book.getSheetAt(0));
		Assert.assertEquals(sheet3, book.getSheetAt(2));
		
		book.moveSheetTo(sheet1, 2);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(2));
		Assert.assertEquals(sheet2, book.getSheetAt(0));
		Assert.assertEquals(sheet3, book.getSheetAt(1));
		
		
		book.moveSheetTo(sheet1, 1);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(1));
		Assert.assertEquals(sheet2, book.getSheetAt(0));
		Assert.assertEquals(sheet3, book.getSheetAt(2));
		
		book.moveSheetTo(sheet1, 0);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheetAt(0));
		Assert.assertEquals(sheet2, book.getSheetAt(1));
		Assert.assertEquals(sheet3, book.getSheetAt(2));
		
		try{
		book.moveSheetTo(sheet1, 3);
		}catch(InvalidateModelOpException x){}//ownership
		
	}
	@Test
	public void testNormal(){
		NBook book = new BookImpl();
		book.createSheet("Sheet1");
		Assert.assertEquals(1, book.getNumOfSheet());
		NSheet sheet = book.createSheet("Sheet2");;
		Assert.assertEquals(-1, sheet.getStartRow());
		Assert.assertEquals(-1, sheet.getEndRow());
		Assert.assertEquals(-1, sheet.getStartColumn());
		Assert.assertEquals(-1, sheet.getEndColumn());
		Assert.assertEquals(-1, sheet.getStartColumn(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		
		NRow row = sheet.getRowAt(3);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(-1, row.getStartColumn());
		Assert.assertEquals(-1, row.getEndColumn());
		NColumn column = sheet.getColumnAt(6);
		Assert.assertEquals(true, column.isNull());
		
		NCell cell = sheet.getCellAt(3, 6);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(3,6)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(3,6)", cell.getValue());
		
		Assert.assertEquals(3, sheet.getStartRow());
		Assert.assertEquals(3, sheet.getEndRow());
		Assert.assertEquals(6, sheet.getStartColumn());
		Assert.assertEquals(6, sheet.getEndColumn());
		
		Assert.assertEquals(-1, sheet.getStartColumn(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumn(3));
		Assert.assertEquals(6, sheet.getEndColumn(3));
		Assert.assertEquals(6, row.getStartColumn());
		Assert.assertEquals(6, row.getEndColumn());
		Assert.assertEquals(-1, sheet.getStartColumn(4));
		Assert.assertEquals(-1, sheet.getEndColumn(4));
		
		
		//another cell
		column = sheet.getColumnAt(12);
		Assert.assertEquals(true, column.isNull());
		
		cell = sheet.getCellAt(3, 12);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(3,12)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(3,12)", cell.getValue());
		
		Assert.assertEquals(6, row.getStartColumn());
		Assert.assertEquals(12, row.getEndColumn());
		
		Assert.assertEquals(3, sheet.getStartRow());
		Assert.assertEquals(3, sheet.getEndRow());
		Assert.assertEquals(6, sheet.getStartColumn());
		Assert.assertEquals(12, sheet.getEndColumn());
		Assert.assertEquals(-1, sheet.getStartColumn(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumn(3));
		Assert.assertEquals(12, sheet.getEndColumn(3));
		Assert.assertEquals(-1, sheet.getStartColumn(4));
		Assert.assertEquals(-1, sheet.getEndColumn(4));
		
		
		//another cell
		row = sheet.getRowAt(4);
		column = sheet.getColumnAt(8);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(true, column.isNull());
		
		cell = sheet.getCellAt(4, 8);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(4,8)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(4,8)", cell.getValue());
		
		Assert.assertEquals(8, row.getStartColumn());
		Assert.assertEquals(8, row.getEndColumn());
		
		Assert.assertEquals(3, sheet.getStartRow());
		Assert.assertEquals(4, sheet.getEndRow());
		Assert.assertEquals(6, sheet.getStartColumn());
		Assert.assertEquals(12, sheet.getEndColumn());
		Assert.assertEquals(-1, sheet.getStartColumn(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumn(3));
		Assert.assertEquals(12, sheet.getEndColumn(3));
		Assert.assertEquals(8, sheet.getStartColumn(4));
		Assert.assertEquals(8, sheet.getEndColumn(4));
		
		
		//another cell
		row = sheet.getRowAt(0);
		column = sheet.getColumnAt(0);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(true, column.isNull());
		
		cell = sheet.getCellAt(0, 0);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(0,0)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(0,0)", cell.getValue());
		
		Assert.assertEquals(0, row.getStartColumn());
		Assert.assertEquals(0, row.getEndColumn());
		
		Assert.assertEquals(0, sheet.getStartRow());
		Assert.assertEquals(4, sheet.getEndRow());
		Assert.assertEquals(0, sheet.getStartColumn());
		Assert.assertEquals(12, sheet.getEndColumn());
		Assert.assertEquals(0, sheet.getStartColumn(0));
		Assert.assertEquals(0, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumn(3));
		Assert.assertEquals(12, sheet.getEndColumn(3));
		Assert.assertEquals(8, sheet.getStartColumn(4));
		Assert.assertEquals(8, sheet.getEndColumn(4));	
		
		StringBuilder builder = new StringBuilder();
		((BookImpl)book).dump(builder);
		System.out.println(builder.toString());
		
	}
}
