package org.zkoss.zss.ngmodel;

import java.util.Iterator;

import junit.framework.Assert;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.ngmodel.impl.TreeMapDataGridImpl;

//this is just poc, not all test can pass
@Ignore
public class TreeMapDataGridTest extends ModelTest{
	protected NSheet initialDataGrid(NSheet sheet){
		sheet.setDataGrid(new TreeMapDataGridImpl());
		return sheet;
	}
	
	@Test
	public void testDataGridIterate(){
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell(10, 12).setValue("ABC");
		sheet1.getCell(23, 6).setValue("DEF");
		sheet1.getCell(23, 2).setValue("123");
		sheet1.getCell(14, 5).setValue("KKK");
		
		NDataGrid grid = sheet1.getDataGrid();
		
		Assert.assertTrue(grid.isProvidedIterator());
		
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
	
	@Test
	public void testSheetIterate(){
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		//set data on data grid, but iterate on sheet
		NDataGrid dg = sheet1.getDataGrid();
		dg.setValue(10, 12, new NCellValue("ABC"));
		dg.setValue(23, 6, new NCellValue("DEF"));
		dg.setValue(23, 2, new NCellValue("123"));
		dg.setValue(14, 5, new NCellValue("KKK"));
		
		
		Iterator<NRow> rows = sheet1.getRowIterator();
		
		NRow row = rows.next();
		Assert.assertEquals(10,row.getIndex());
		
		Iterator<NCell> cells = sheet1.getCellIterator(row.getIndex());
		NCell cell = cells.next();
		Assert.assertEquals(10,cell.getRowIndex());
		Assert.assertEquals(12,cell.getColumnIndex());
		Assert.assertEquals("ABC", cell.getValue());
		
		Assert.assertFalse(cells.hasNext());
		
		
		row = rows.next();
		Assert.assertEquals(14,row.getIndex());
		
		cells = sheet1.getCellIterator(row.getIndex());
		cell = cells.next();
		Assert.assertEquals(14,cell.getRowIndex());
		Assert.assertEquals(5,cell.getColumnIndex());
		Assert.assertEquals("KKK", cell.getValue());
		
		row = rows.next();
		Assert.assertEquals(23,row.getIndex());
		
		cells = sheet1.getCellIterator(row.getIndex());
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(2,cell.getColumnIndex());
		Assert.assertEquals("123", cell.getValue());
		
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(6,cell.getColumnIndex());
		Assert.assertEquals("DEF", cell.getValue());
		
		Assert.assertFalse(cells.hasNext());
		Assert.assertFalse(rows.hasNext());
	}
	
	@Test
	public void testSheetDataMixedIterate(){
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		//set data on data grid, but iterate on sheet
		NDataGrid dg = sheet1.getDataGrid();
		dg.setValue(10, 12, new NCellValue("ABC"));
		dg.setValue(23, 6, new NCellValue("DEF"));
		sheet1.getCell(23, 2).setValue("123");
		sheet1.getCell(14, 5).setValue("KKK");
		
		
		Iterator<NRow> rows = sheet1.getRowIterator();
		
		NRow row = rows.next();
		Assert.assertEquals(10,row.getIndex());
		
		Iterator<NCell> cells = sheet1.getCellIterator(row.getIndex());
		NCell cell = cells.next();
		Assert.assertEquals(10,cell.getRowIndex());
		Assert.assertEquals(12,cell.getColumnIndex());
		Assert.assertEquals("ABC", cell.getValue());
		
		Assert.assertFalse(cells.hasNext());
		
		
		row = rows.next();
		Assert.assertEquals(14,row.getIndex());
		
		cells = sheet1.getCellIterator(row.getIndex());
		cell = cells.next();
		Assert.assertEquals(14,cell.getRowIndex());
		Assert.assertEquals(5,cell.getColumnIndex());
		Assert.assertEquals("KKK", cell.getValue());
		
		row = rows.next();
		Assert.assertEquals(23,row.getIndex());
		
		cells = sheet1.getCellIterator(row.getIndex());
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(2,cell.getColumnIndex());
		Assert.assertEquals("123", cell.getValue());
		
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(6,cell.getColumnIndex());
		Assert.assertEquals("DEF", cell.getValue());
		
		Assert.assertFalse(cells.hasNext());
		Assert.assertFalse(rows.hasNext());
	}
	
	@Test
	public void testSheetDataMixedIterate2(){
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		//set data on data grid, but iterate on sheet
		NDataGrid dg = sheet1.getDataGrid();
		sheet1.getCell(10, 12).setValue("ABC");
		sheet1.getCell(23, 6).setValue("DEF");
		dg.setValue(23, 2, new NCellValue("123"));
		dg.setValue(14, 5, new NCellValue("KKK"));
		
		
		Iterator<NRow> rows = sheet1.getRowIterator();
		
		NRow row = rows.next();
		Assert.assertEquals(10,row.getIndex());
		
		Iterator<NCell> cells = sheet1.getCellIterator(row.getIndex());
		NCell cell = cells.next();
		Assert.assertEquals(10,cell.getRowIndex());
		Assert.assertEquals(12,cell.getColumnIndex());
		Assert.assertEquals("ABC", cell.getValue());
		
		Assert.assertFalse(cells.hasNext());
		
		
		row = rows.next();
		Assert.assertEquals(14,row.getIndex());
		
		cells = sheet1.getCellIterator(row.getIndex());
		cell = cells.next();
		Assert.assertEquals(14,cell.getRowIndex());
		Assert.assertEquals(5,cell.getColumnIndex());
		Assert.assertEquals("KKK", cell.getValue());
		
		row = rows.next();
		Assert.assertEquals(23,row.getIndex());
		
		cells = sheet1.getCellIterator(row.getIndex());
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(2,cell.getColumnIndex());
		Assert.assertEquals("123", cell.getValue());
		
		cell = cells.next();
		Assert.assertEquals(23,cell.getRowIndex());
		Assert.assertEquals(6,cell.getColumnIndex());
		Assert.assertEquals("DEF", cell.getValue());
		
		Assert.assertFalse(cells.hasNext());
		Assert.assertFalse(rows.hasNext());
	}
	
	@Test
	public void testSheetDataMixedIterate3(){
		NBook book = NBooks.createBook("book1");
		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		//set data on data grid, but iterate on sheet
		NDataGrid dg = sheet1.getDataGrid();
		sheet1.getCell(10, 12).setCellStyle(book.createCellStyle(false));//update style only
		sheet1.getCell(23, 6).setCellStyle(book.createCellStyle(false));//update style only
		dg.setValue(23, 2, new NCellValue("123"));
		dg.setValue(14, 5, new NCellValue("KKK"));
		
		
		NDataGrid grid = sheet1.getDataGrid();
		
		Assert.assertTrue(grid.isProvidedIterator());
		
		Iterator<NDataRow> rows = grid.getRowIterator();
		
		Iterator<NDataCell> cells;
		NDataCell cell;
		NDataRow row;
		
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
		
		Assert.assertFalse(cells.hasNext());
		Assert.assertFalse(rows.hasNext());
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
