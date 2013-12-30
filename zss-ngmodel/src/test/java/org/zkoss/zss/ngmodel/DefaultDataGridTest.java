package org.zkoss.zss.ngmodel;

import junit.framework.Assert;

import org.junit.Test;

public class DefaultDataGridTest extends ModelTest{
	protected NSheet initialDataGrid(NSheet sheet){
		sheet.setDataGrid(new DefaultDataGrid(sheet));
		return sheet;
	}
	
	@Test
	public void testDataGridStartEnd(){
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		Assert.assertEquals(-1,sheet1.getStartRowIndex());
		Assert.assertEquals(-1,sheet1.getEndRowIndex());
		
		NDataGrid grid = sheet1.getDataGrid();
		
		sheet1.getCell(10, 2).setValue("ABC");
		sheet1.getCell(10, 12).setValue("ABC");
		sheet1.getCell(23, 6).setValue("DEF");
		sheet1.getCell(23, 2).setValue("123");
		sheet1.getCell(15, 3).setValue("ABC");
		
		Assert.assertEquals(10,sheet1.getStartRowIndex());
		Assert.assertEquals(23,sheet1.getEndRowIndex());
		
		Assert.assertEquals(2,sheet1.getStartCellIndex(10));
		Assert.assertEquals(12,sheet1.getEndCellIndex(10));
		
		Assert.assertEquals(3,sheet1.getStartCellIndex(15));
		Assert.assertEquals(3,sheet1.getEndCellIndex(15));
		
		Assert.assertEquals(2,sheet1.getStartCellIndex(23));
		Assert.assertEquals(6,sheet1.getEndCellIndex(23));
		
		Assert.assertEquals(-1,sheet1.getStartCellIndex(3));
		Assert.assertEquals(-1,sheet1.getEndCellIndex(3));
		
		grid.setValue(3, 1, new NCellValue("ABC"));
		
		grid.setValue(15, 4, new NCellValue("ABC"));
	
		grid.setValue(23, 1, new NCellValue("ABC"));
		
		Assert.assertEquals(3,sheet1.getStartRowIndex());
		Assert.assertEquals(23,sheet1.getEndRowIndex());
		
		Assert.assertEquals(2,sheet1.getStartCellIndex(10));
		Assert.assertEquals(12,sheet1.getEndCellIndex(10));
		
		Assert.assertEquals(3,sheet1.getStartCellIndex(15));
		Assert.assertEquals(4,sheet1.getEndCellIndex(15));
		
		Assert.assertEquals(1,sheet1.getStartCellIndex(23));
		Assert.assertEquals(6,sheet1.getEndCellIndex(23));
		
		Assert.assertEquals(1,sheet1.getStartCellIndex(3));
		Assert.assertEquals(1,sheet1.getEndCellIndex(3));
	}
}
