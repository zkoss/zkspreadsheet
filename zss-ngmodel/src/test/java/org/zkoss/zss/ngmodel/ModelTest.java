package org.zkoss.zss.ngmodel;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class ModelTest {

	
	@Test
	public void testSheet(){
		NBook book = new BookImpl();
		Assert.assertEquals(1, book.getNumOfSheet());
		NSheet sheet = book.getSheetAt(0);
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
		
		cell.setValue("At(3,6)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("At(3,6)", cell.getValue());
		
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
		
		cell.setValue("At(3,12)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("At(3,12)", cell.getValue());
		
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
		
		cell.setValue("At(4,8)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("At(4,8)", cell.getValue());
		
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
		
		cell.setValue("At(0,0)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("At(0,0)", cell.getValue());
		
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
