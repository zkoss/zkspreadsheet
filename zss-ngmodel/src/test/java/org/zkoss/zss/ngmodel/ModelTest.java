package org.zkoss.zss.ngmodel;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class ModelTest {

	@Test
	public void testSheet(){
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
		Assert.assertEquals(sheet1, book.getSheet(0));
		Assert.assertEquals(sheet2, book.getSheet(1));
		Assert.assertEquals(sheet1, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(null, book.getSheetByName("Sheet3"));
		
		book.deleteSheet(sheet1);
		
		Assert.assertEquals(1, book.getNumOfSheet());
		Assert.assertEquals(sheet2, book.getSheet(0));
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
		Assert.assertEquals(sheet1, book.getSheet(1));
		Assert.assertEquals(sheet2, book.getSheet(0));
		Assert.assertEquals(sheet1, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(null, book.getSheetByName("Sheet3"));
		
		NSheet sheet3 = book.createSheet("Sheet3");
		
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheet(1));
		Assert.assertEquals(sheet2, book.getSheet(0));
		Assert.assertEquals(sheet3, book.getSheet(2));
		Assert.assertEquals(sheet1, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(sheet3, book.getSheetByName("Sheet3"));
		
		
		book.moveSheetTo(sheet1, 0);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheet(0));
		Assert.assertEquals(sheet2, book.getSheet(1));
		Assert.assertEquals(sheet3, book.getSheet(2));
		
		book.moveSheetTo(sheet1, 1);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheet(1));
		Assert.assertEquals(sheet2, book.getSheet(0));
		Assert.assertEquals(sheet3, book.getSheet(2));
		
		book.moveSheetTo(sheet1, 2);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheet(2));
		Assert.assertEquals(sheet2, book.getSheet(0));
		Assert.assertEquals(sheet3, book.getSheet(1));
		
		
		book.moveSheetTo(sheet1, 1);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheet(1));
		Assert.assertEquals(sheet2, book.getSheet(0));
		Assert.assertEquals(sheet3, book.getSheet(2));
		
		book.moveSheetTo(sheet1, 0);
		Assert.assertEquals(3, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheet(0));
		Assert.assertEquals(sheet2, book.getSheet(1));
		Assert.assertEquals(sheet3, book.getSheet(2));
		
		try{
		book.moveSheetTo(sheet1, 3);
		}catch(InvalidateModelOpException x){}//ownership
		
	}
	@Test
	public void testCellRange(){
		NBook book = new BookImpl();
		book.createSheet("Sheet1");
		Assert.assertEquals(1, book.getNumOfSheet());
		NSheet sheet = book.createSheet("Sheet2");;
		Assert.assertEquals(-1, sheet.getStartRowIndex());
		Assert.assertEquals(-1, sheet.getEndRowIndex());
		Assert.assertEquals(-1, sheet.getStartColumnIndex());
		Assert.assertEquals(-1, sheet.getEndColumnIndex());
		Assert.assertEquals(-1, sheet.getStartColumnIndex(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		
		NRow row = sheet.getRow(3);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(-1, row.getStartCellIndex());
		Assert.assertEquals(-1, row.getEndCellIndex());
		NColumn column = sheet.getColumn(6);
		Assert.assertEquals(true, column.isNull());
		
		NCell cell = sheet.getCell(3, 6);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(3,6)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(3,6)", cell.getValue());
		
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(3, sheet.getEndRowIndex());
		Assert.assertEquals(6, sheet.getStartColumnIndex());
		Assert.assertEquals(6, sheet.getEndColumnIndex());
		
		Assert.assertEquals(-1, sheet.getStartColumnIndex(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumnIndex(3));
		Assert.assertEquals(6, sheet.getEndColumn(3));
		Assert.assertEquals(6, row.getStartCellIndex());
		Assert.assertEquals(6, row.getEndCellIndex());
		Assert.assertEquals(-1, sheet.getStartColumnIndex(4));
		Assert.assertEquals(-1, sheet.getEndColumn(4));
		
		
		//another cell
		column = sheet.getColumn(12);
		Assert.assertEquals(true, column.isNull());
		
		cell = sheet.getCell(3, 12);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(3,12)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(3,12)", cell.getValue());
		
		Assert.assertEquals(6, row.getStartCellIndex());
		Assert.assertEquals(12, row.getEndCellIndex());
		
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(3, sheet.getEndRowIndex());
		Assert.assertEquals(6, sheet.getStartColumnIndex());
		Assert.assertEquals(12, sheet.getEndColumnIndex());
		Assert.assertEquals(-1, sheet.getStartColumnIndex(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumnIndex(3));
		Assert.assertEquals(12, sheet.getEndColumn(3));
		Assert.assertEquals(-1, sheet.getStartColumnIndex(4));
		Assert.assertEquals(-1, sheet.getEndColumn(4));
		
		
		//another cell
		row = sheet.getRow(4);
		column = sheet.getColumn(8);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(true, column.isNull());
		
		cell = sheet.getCell(4, 8);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(4,8)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(4,8)", cell.getValue());
		
		Assert.assertEquals(8, row.getStartCellIndex());
		Assert.assertEquals(8, row.getEndCellIndex());
		
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(4, sheet.getEndRowIndex());
		Assert.assertEquals(6, sheet.getStartColumnIndex());
		Assert.assertEquals(12, sheet.getEndColumnIndex());
		Assert.assertEquals(-1, sheet.getStartColumnIndex(0));
		Assert.assertEquals(-1, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumnIndex(3));
		Assert.assertEquals(12, sheet.getEndColumn(3));
		Assert.assertEquals(8, sheet.getStartColumnIndex(4));
		Assert.assertEquals(8, sheet.getEndColumn(4));
		
		
		//another cell
		row = sheet.getRow(0);
		column = sheet.getColumn(0);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(true, column.isNull());
		
		cell = sheet.getCell(0, 0);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(0,0)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(0,0)", cell.getValue());
		
		Assert.assertEquals(0, row.getStartCellIndex());
		Assert.assertEquals(0, row.getEndCellIndex());
		
		Assert.assertEquals(0, sheet.getStartRowIndex());
		Assert.assertEquals(4, sheet.getEndRowIndex());
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(12, sheet.getEndColumnIndex());
		Assert.assertEquals(0, sheet.getStartColumnIndex(0));
		Assert.assertEquals(0, sheet.getEndColumn(0));
		Assert.assertEquals(6, sheet.getStartColumnIndex(3));
		Assert.assertEquals(12, sheet.getEndColumn(3));
		Assert.assertEquals(8, sheet.getStartColumnIndex(4));
		Assert.assertEquals(8, sheet.getEndColumn(4));	
	}
	
	@Test
	public void testClearSheetRow(){
		NBook book = new BookImpl();
		NSheet sheet = book.createSheet("Sheet 1");
		
		for(int i=10;i<=20;i+=2){
			for(int j=3;j<=15;j+=3){
				NCell cell = sheet.getCell(i, j);
				cell.setValue("("+i+","+j+")");
			}
		}
		Assert.assertEquals(false, sheet.getRow(10).isNull());
		Assert.assertEquals(false, sheet.getRow(12).isNull());
		Assert.assertEquals(false, sheet.getRow(14).isNull());
		Assert.assertEquals(false, sheet.getRow(16).isNull());
		
		Assert.assertEquals("(10,3)", sheet.getCell(10, 3).getValue());
		Assert.assertEquals("(12,6)", sheet.getCell(12, 6).getValue());
		Assert.assertEquals("(14,9)", sheet.getCell(14, 9).getValue());
		Assert.assertEquals("(16,12)", sheet.getCell(16, 12).getValue());
		
		sheet.clearRow(12, 14);
		
		Assert.assertEquals(false, sheet.getRow(10).isNull());
		Assert.assertEquals(true, sheet.getRow(12).isNull());
		Assert.assertEquals(true, sheet.getRow(14).isNull());
		Assert.assertEquals(false, sheet.getRow(16).isNull());
		
		Assert.assertEquals("(10,3)", sheet.getCell(10, 3).getValue());
		Assert.assertEquals(null, sheet.getCell(12, 6).getValue());
		Assert.assertEquals(null, sheet.getCell(14, 9).getValue());
		Assert.assertEquals(true, sheet.getCell(12, 6).isNull());
		Assert.assertEquals(true, sheet.getCell(14, 9).isNull());
		Assert.assertEquals("(16,12)", sheet.getCell(16, 12).getValue());
		
		Assert.assertEquals(10, sheet.getStartRowIndex());
		Assert.assertEquals(20, sheet.getEndRowIndex());
		
		Assert.assertEquals(3, sheet.getStartColumnIndex());
		Assert.assertEquals(15, sheet.getEndColumnIndex());
		
		
		sheet.clearRow(1, 100);
		Assert.assertEquals(true, sheet.getRow(10).isNull());
		Assert.assertEquals(true, sheet.getRow(12).isNull());
		Assert.assertEquals(true, sheet.getRow(14).isNull());
		Assert.assertEquals(true, sheet.getRow(16).isNull());
		Assert.assertEquals(true, sheet.getRow(18).isNull());
		Assert.assertEquals(true, sheet.getRow(20).isNull());
		
		Assert.assertEquals(-1, sheet.getStartRowIndex());
		Assert.assertEquals(-1, sheet.getEndRowIndex());
		
		//clear all down doesn't effect the existed column status.(think about column style on the column)
		Assert.assertEquals(3, sheet.getStartColumnIndex());
		Assert.assertEquals(15, sheet.getEndColumnIndex());
		
	}
	
	@Test
	public void testClearSheetColumn(){
		NBook book = new BookImpl();
		NSheet sheet = book.createSheet("Sheet 1");
		
		for(int i=10;i<=20;i+=2){
			for(int j=3;j<=15;j+=3){
				NCell cell = sheet.getCell(j, i);
				cell.setValue("("+j+","+i+")");
			}
		}
		
		Assert.assertEquals(false, sheet.getColumn(10).isNull());
		Assert.assertEquals(false, sheet.getColumn(12).isNull());
		Assert.assertEquals(false, sheet.getColumn(14).isNull());
		Assert.assertEquals(false, sheet.getColumn(16).isNull());
		
		Assert.assertEquals("(3,10)", sheet.getCell(3,10).getValue());
		Assert.assertEquals("(6,12)", sheet.getCell(6,12).getValue());
		Assert.assertEquals("(9,14)", sheet.getCell(9,14).getValue());
		Assert.assertEquals("(12,16)", sheet.getCell(12,16).getValue());
		
		sheet.clearColumn(12, 14);
		
		
		Assert.assertEquals(false, sheet.getColumn(10).isNull());
		Assert.assertEquals(true, sheet.getColumn(12).isNull());
		Assert.assertEquals(true, sheet.getColumn(14).isNull());
		Assert.assertEquals(false, sheet.getColumn(16).isNull());
		
		Assert.assertEquals("(3,10)", sheet.getCell(3, 10).getValue());
		Assert.assertEquals(null, sheet.getCell(6, 12).getValue());
		Assert.assertEquals(null, sheet.getCell(9, 14).getValue());
		Assert.assertEquals(true, sheet.getCell(6, 12).isNull());
		Assert.assertEquals(true, sheet.getCell(9, 14).isNull());
		Assert.assertEquals("(12,16)", sheet.getCell(12, 16).getValue());
		
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(15, sheet.getEndRowIndex());
		
		Assert.assertEquals(10, sheet.getStartColumnIndex());
		Assert.assertEquals(20, sheet.getEndColumnIndex());
		
		
		sheet.clearColumn(1, 30);
		Assert.assertEquals(true, sheet.getColumn(10).isNull());
		Assert.assertEquals(true, sheet.getColumn(12).isNull());
		Assert.assertEquals(true, sheet.getColumn(14).isNull());
		Assert.assertEquals(true, sheet.getColumn(16).isNull());
		Assert.assertEquals(true, sheet.getColumn(18).isNull());
		Assert.assertEquals(true, sheet.getColumn(20).isNull());
		
		//clear all down doesn't effect the existed column status.(think about column style on the column)
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(15, sheet.getEndRowIndex());
		
		
		Assert.assertEquals(-1, sheet.getStartColumnIndex());
		Assert.assertEquals(-1, sheet.getEndColumnIndex());

	}
	
	@Test
	public void testClearSheetCell(){
		NBook book = new BookImpl();
		NSheet sheet = book.createSheet("Sheet 1");
		
		for(int i=10;i<=20;i+=2){
			for(int j=3;j<=15;j+=3){
				NCell cell = sheet.getCell(i, j);
				cell.setValue("("+i+","+j+")");
			}
		}
		Assert.assertEquals(false, sheet.getRow(10).isNull());
		Assert.assertEquals(false, sheet.getRow(12).isNull());
		Assert.assertEquals(false, sheet.getRow(14).isNull());
		Assert.assertEquals(false, sheet.getRow(16).isNull());
		
		Assert.assertEquals("(10,3)", sheet.getCell(10, 3).getValue());
		Assert.assertEquals("(12,6)", sheet.getCell(12, 6).getValue());
		Assert.assertEquals("(14,9)", sheet.getCell(14, 9).getValue());
		Assert.assertEquals("(16,12)", sheet.getCell(16, 12).getValue());
		
		dump(book);
		
		sheet.clearCell(12, 6 ,14, 9);
		
		Assert.assertEquals(false, sheet.getRow(10).isNull());
		Assert.assertEquals(false, sheet.getRow(12).isNull());
		Assert.assertEquals(false, sheet.getRow(14).isNull());
		Assert.assertEquals(false, sheet.getRow(16).isNull());
		Assert.assertEquals(false, sheet.getColumn(3).isNull());
		Assert.assertEquals(false, sheet.getColumn(6).isNull());
		Assert.assertEquals(false, sheet.getColumn(9).isNull());
		Assert.assertEquals(false, sheet.getColumn(12).isNull());
		
		Assert.assertEquals("(10,3)", sheet.getCell(10, 3).getValue());
		Assert.assertEquals(null, sheet.getCell(12, 6).getValue());
		Assert.assertEquals(null, sheet.getCell(14, 9).getValue());
		Assert.assertEquals(true, sheet.getCell(12, 6).isNull());
		Assert.assertEquals(true, sheet.getCell(14, 9).isNull());
		Assert.assertEquals("(16,12)", sheet.getCell(16, 12).getValue());
		
		Assert.assertEquals(10, sheet.getStartRowIndex());
		Assert.assertEquals(20, sheet.getEndRowIndex());
		
		Assert.assertEquals(3, sheet.getStartColumnIndex());
		Assert.assertEquals(15, sheet.getEndColumnIndex());
		
		
		sheet.clearCell(1, 1 ,100, 50);
		Assert.assertEquals(false, sheet.getRow(10).isNull());
		Assert.assertEquals(false, sheet.getRow(12).isNull());
		Assert.assertEquals(false, sheet.getRow(14).isNull());
		Assert.assertEquals(false, sheet.getRow(16).isNull());
		Assert.assertEquals(false, sheet.getRow(18).isNull());
		Assert.assertEquals(false, sheet.getRow(20).isNull());
		
		Assert.assertEquals(false, sheet.getColumn(3).isNull());
		Assert.assertEquals(false, sheet.getColumn(6).isNull());
		Assert.assertEquals(false, sheet.getColumn(9).isNull());
		Assert.assertEquals(false, sheet.getColumn(12).isNull());
		Assert.assertEquals(false, sheet.getColumn(15).isNull());
		
		Assert.assertEquals(10, sheet.getStartRowIndex());
		Assert.assertEquals(20, sheet.getEndRowIndex());
		
		Assert.assertEquals(3, sheet.getStartColumnIndex());
		Assert.assertEquals(15, sheet.getEndColumnIndex());
		
		dump(book);
		
	}
	
	public static void dump(NBook book){
		StringBuilder builder = new StringBuilder();
		((BookImpl)book).dump(builder);
		System.out.println(builder.toString());
	}
}
