package org.zkoss.zss.ngmodel;

import java.util.Iterator;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.zss.ngmodel.impl.TreeMapDataGridImpl;

public class ModelTreeMapDataGridTest extends ModelTest{
	protected NSheet initialDataGrid(NSheet sheet){
		sheet.setDataGrid(new TreeMapDataGridImpl());
		return sheet;
	}
	
	@Test
	public void testIterate(){
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell(10, 12).setValue("ABC");
		sheet1.getCell(23, 6).setValue("DEF");
		sheet1.getCell(23, 2).setValue("123");
		sheet1.getCell(14, 5).setValue("KKK");
		
		NDataGrid grid = sheet1.getDataGrid();
		
		Assert.assertTrue(grid.supportDataIterator());
		
		Iterator<NDataRow> rows = grid.getRowIterator();
		
		NDataRow row = rows.next();
		Assert.assertEquals(10,row.getIndex());
		
		Iterator<NDataCell> cells = row.getCellIterator();
		NDataCell cell = cells.next();
		Assert.assertEquals(10,cell.getRowIndex());
		Assert.assertEquals(12,cell.getColumnIndex());
		Assert.assertEquals("ABC", cell.getValue().getValue());
		
		Assert.assertFalse(cells.hasNext());
		
		
		row = rows.next();
		Assert.assertEquals(14,row.getIndex());
		
		cells = row.getCellIterator();
		cell = cells.next();
		Assert.assertEquals(14,cell.getRowIndex());
		Assert.assertEquals(5,cell.getColumnIndex());
		Assert.assertEquals("KKK", cell.getValue().getValue());
		
		row = rows.next();
		Assert.assertEquals(23,row.getIndex());
		
		cells = row.getCellIterator();
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(2,cell.getColumnIndex());
		Assert.assertEquals("123", cell.getValue().getValue());
		
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(6,cell.getColumnIndex());
		Assert.assertEquals("DEF", cell.getValue().getValue());
		
		Assert.assertFalse(cells.hasNext());
		Assert.assertFalse(rows.hasNext());
	}
}
