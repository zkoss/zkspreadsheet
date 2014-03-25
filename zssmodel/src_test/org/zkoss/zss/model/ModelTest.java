package org.zkoss.zss.model;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.concurrent.locks.ReadWriteLock;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SChart;
import org.zkoss.zss.model.SColumn;
import org.zkoss.zss.model.SColumnArray;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SName;
import org.zkoss.zss.model.SPicture;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.ViewAnchor;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.SAutoFilter.FilterOp;
import org.zkoss.zss.model.SAutoFilter.NFilterColumn;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SCellStyle.Alignment;
import org.zkoss.zss.model.SCellStyle.BorderType;
import org.zkoss.zss.model.SCellStyle.FillPattern;
import org.zkoss.zss.model.SChart.ChartType;
import org.zkoss.zss.model.SDataValidation.ValidationType;
import org.zkoss.zss.model.SFont.Boldweight;
import org.zkoss.zss.model.SFont.TypeOffset;
import org.zkoss.zss.model.SFont.Underline;
import org.zkoss.zss.model.SHyperlink.HyperlinkType;
import org.zkoss.zss.model.SPicture.Format;
import org.zkoss.zss.model.chart.SGeneralChartData;
import org.zkoss.zss.model.chart.SSeries;
import org.zkoss.zss.model.impl.AbstractBookSeriesAdv;
import org.zkoss.zss.model.impl.AbstractCellAdv;
import org.zkoss.zss.model.impl.AbstractSheetAdv;
import org.zkoss.zss.model.impl.BookImpl;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.impl.SheetImpl;
import org.zkoss.zss.model.impl.chart.GeneralChartDataImpl;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.util.CellStyleMatcher;
import org.zkoss.zss.model.util.FontMatcher;
import org.zkoss.zss.model.util.Validations;
import org.zkoss.zss.range.SRange.DeleteShift;
import org.zkoss.zss.range.SRange.InsertCopyOrigin;
import org.zkoss.zss.range.SRange.InsertShift;
import org.zkoss.zss.range.SRanges;

public class ModelTest {

	
	@BeforeClass
	static public void beforeClass() {
		Setup.touch();
	}
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	protected SSheet initialDataGrid(SSheet sheet){
		return sheet;
	}

	
	@Test 
	public void testLock(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		Assert.assertEquals(1, book.getNumOfSheet());
		SSheet sheet2 = initialDataGrid(book.createSheet("Sheet2"));
		Assert.assertEquals(2, book.getNumOfSheet());
		
		ReadWriteLock l = book.getBookSeries().getLock();
		
		//Write Read
		System.out.println("A");
		l.writeLock().lock();
		try{
			System.out.println("B");
			l.readLock().lock();
			System.out.println("C");
			
			l.readLock().unlock();
			System.out.println("D");
		}finally{
			System.out.println("E");
			l.writeLock().unlock();
		}
		System.out.println("F");
		System.out.println("End Write Read");
		
		//Write Write
		System.out.println("A");
		l.writeLock().lock();
		try{
			System.out.println("B");
			l.writeLock().lock();
			System.out.println("C");
			
			l.writeLock().unlock();
			System.out.println("D");
		}finally{
			System.out.println("E");
			l.writeLock().unlock();
		}
		System.out.println("F");
		System.out.println("End Write Write");
		
		//Read Read
		System.out.println("A");
		l.readLock().lock();
		try{
			System.out.println("B");
			l.readLock().lock();
			System.out.println("C");
			
			l.readLock().unlock();
			System.out.println("D");
		}finally{
			System.out.println("E");
			l.readLock().unlock();
		}
		System.out.println("F");
		System.out.println("End Read Read");
			
		
//		System.out.println("A");
//		l.readLock().lock();
//		try{
//			System.out.println("B");
//			l.writeLock().lock();
//			System.out.println("C");
//			
//			l.writeLock().unlock();
//			System.out.println("D");
//		}finally{
//			System.out.println("E");
//			l.readLock().unlock();
//		}
//		System.out.println("F");
//		System.out.println("End Read Write");		
		
		//Write Write
		System.out.println("A");
		l.writeLock().lock();
		System.out.println("A2");
		l.readLock().lock();
		try{
			System.out.println("B");
			l.writeLock().lock();
			System.out.println("C");
			
			l.writeLock().unlock();
			System.out.println("D");
		}finally{
			System.out.println("E2");
			l.readLock().unlock();
			System.out.println("E");
			l.writeLock().unlock();
		}
		System.out.println("F");
		System.out.println("End Write Read Write");		
	}

	
	@Test 
	public void testRegion(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
	
		CellRegion region = new CellRegion("A1:B3");
		
		Assert.assertEquals("A1:B3", region.getReferenceString());
		Assert.assertEquals(0, region.getRow());
		Assert.assertEquals(0, region.getColumn());
		Assert.assertEquals(2, region.getLastRow());
		Assert.assertEquals(1, region.getLastColumn());
		
		
		region = new CellRegion("B3:A1");
		Assert.assertEquals("A1:B3", region.getReferenceString());
		Assert.assertEquals(0, region.getRow());
		Assert.assertEquals(0, region.getColumn());
		Assert.assertEquals(2, region.getLastRow());
		Assert.assertEquals(1, region.getLastColumn());
		
		region = new CellRegion("B1:A3");
		Assert.assertEquals("A1:B3", region.getReferenceString());
		Assert.assertEquals(0, region.getRow());
		Assert.assertEquals(0, region.getColumn());
		Assert.assertEquals(2, region.getLastRow());
		Assert.assertEquals(1, region.getLastColumn());
		
		region = new CellRegion("A3:B1");
		Assert.assertEquals("A1:B3", region.getReferenceString());
		Assert.assertEquals(0, region.getRow());
		Assert.assertEquals(0, region.getColumn());
		Assert.assertEquals(2, region.getLastRow());
		Assert.assertEquals(1, region.getLastColumn());
		
		
		SheetRegion sregion = new SheetRegion(sheet1,"A1:B3");
		
		Assert.assertEquals("Sheet1!A1:B3", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(0, sregion.getColumn());
		Assert.assertEquals(2, sregion.getLastRow());
		Assert.assertEquals(1, sregion.getLastColumn());
		
		
		sregion = new SheetRegion(sheet1,"B3:A1");
		Assert.assertEquals("Sheet1!A1:B3", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(0, sregion.getColumn());
		Assert.assertEquals(2, sregion.getLastRow());
		Assert.assertEquals(1, sregion.getLastColumn());
		
		sregion = new SheetRegion(sheet1,"B1:A3");
		Assert.assertEquals("Sheet1!A1:B3", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(0, sregion.getColumn());
		Assert.assertEquals(2, sregion.getLastRow());
		Assert.assertEquals(1, sregion.getLastColumn());
		
		sregion = new SheetRegion(sheet1,"A3:B1");
		Assert.assertEquals("Sheet1!A1:B3", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(0, sregion.getColumn());
		Assert.assertEquals(2, sregion.getLastRow());
		Assert.assertEquals(1, sregion.getLastColumn());
		
		
		sregion = new SheetRegion(sheet1,"1:3");
		Assert.assertEquals("Sheet1!A1:XFD3", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(0, sregion.getColumn());
		Assert.assertEquals(2, sregion.getLastRow());
		Assert.assertEquals(sheet1.getBook().getMaxColumnIndex(), sregion.getLastColumn());
		
		sregion = new SheetRegion(sheet1,"3:1");
		Assert.assertEquals("Sheet1!A1:XFD3", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(0, sregion.getColumn());
		Assert.assertEquals(2, sregion.getLastRow());
		Assert.assertEquals(sheet1.getBook().getMaxColumnIndex(), sregion.getLastColumn());
		
		
		sregion = new SheetRegion(sheet1,"B:C");
		
		//Too bad, we poi's implementation always use 65535 to be the maxRow Number 
		Assert.assertEquals("Sheet1!B1:C1048576", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(1, sregion.getColumn());
		Assert.assertEquals(sheet1.getBook().getMaxRowIndex(), sregion.getLastRow());
		Assert.assertEquals(2, sregion.getLastColumn());
		
		
		sregion = new SheetRegion(sheet1,"C:B");
		Assert.assertEquals("Sheet1!B1:C1048576", sregion.getReferenceString());
		Assert.assertEquals(0, sregion.getRow());
		Assert.assertEquals(1, sregion.getColumn());
		Assert.assertEquals(sheet1.getBook().getMaxRowIndex(), sregion.getLastRow());
		Assert.assertEquals(2, sregion.getLastColumn());
	}
	
	@Test
	public void testSheet(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		Assert.assertEquals(1, book.getNumOfSheet());
		SSheet sheet2 = initialDataGrid(book.createSheet("Sheet2"));
		Assert.assertEquals(2, book.getNumOfSheet());
		
		try{
			SSheet sheet = initialDataGrid(book.createSheet("Sheet2"));
			Assert.fail("should get exception");
		}catch(InvalidModelOpException x){}
		
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
		}catch(IllegalStateException x){}//ownership
		
		try{
			initialDataGrid(book.createSheet("Sheet3", sheet1));
			Assert.fail("should get exception");
		}catch(IllegalStateException x){}//ownership
		
		try{
			book.moveSheetTo(sheet1, 0);
			Assert.fail("should get exception");
		}catch(IllegalStateException x){}//ownership
		
		sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		Assert.assertEquals(2, book.getNumOfSheet());
		Assert.assertEquals(sheet1, book.getSheet(1));
		Assert.assertEquals(sheet2, book.getSheet(0));
		Assert.assertEquals(sheet1, book.getSheetByName("Sheet1"));
		Assert.assertEquals(sheet2, book.getSheetByName("Sheet2"));
		Assert.assertEquals(null, book.getSheetByName("Sheet3"));
		
		SSheet sheet3 = initialDataGrid(book.createSheet("Sheet3"));
		
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
		}catch(InvalidModelOpException x){}//ownership
		
		
		
	}
	@Test
	public void testSetupColumnArray(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		Assert.assertNull(sheet1.getColumnArray(0));
		
		sheet1.setupColumnArray(0, 3).setWidth(10);
		sheet1.setupColumnArray(4, 5);
		arrays = sheet1.getColumnArrayIterator();
		SColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(3, array.getLastIndex());
		Assert.assertEquals(10, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(4, array.getIndex());
		Assert.assertEquals(5, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		Assert.assertFalse(arrays.hasNext());
		try{
			sheet1.setupColumnArray(5,6);
			Assert.fail();
		}catch(IllegalStateException x){}
		
		sheet1.setupColumnArray(10, 20).setWidth(40);
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(3, array.getLastIndex());
		Assert.assertEquals(10, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(4, array.getIndex());
		Assert.assertEquals(5, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(6, array.getIndex());
		Assert.assertEquals(9, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(10, array.getIndex());
		Assert.assertEquals(20, array.getLastIndex());
		Assert.assertEquals(40, array.getWidth());
		Assert.assertFalse(arrays.hasNext());
	}

	@Test
	public void testSetupColumnArray2() {
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));

		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());

		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());

		Assert.assertNull(sheet1.getColumnArray(0));

		sheet1.setupColumnArray(1, 1).setWidth(10);
		arrays = sheet1.getColumnArrayIterator();
		SColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(0, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());

		array = arrays.next();
		Assert.assertEquals(1, array.getIndex());
		Assert.assertEquals(1, array.getLastIndex());
		Assert.assertEquals(10, array.getWidth());

		Assert.assertFalse(arrays.hasNext());

	}
	
	@Test
	public void testColumnArray(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		
		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		Assert.assertNull(sheet1.getColumnArray(0));
		SColumn column = sheet1.getColumn(0);
		Assert.assertNull(sheet1.getColumnArray(0));
		column.setWidth(150);
		SColumnArray array;
		Assert.assertNotNull(array = sheet1.getColumnArray(0));
		Assert.assertEquals(150, array.getWidth());
		
		Assert.assertEquals(false, array.isHidden());
		column.setHidden(true);
		Assert.assertEquals(true, array.isHidden());
		
		SCellStyle style;
		Assert.assertEquals(book.getDefaultCellStyle(), array.getCellStyle());
		column.setCellStyle(style = book.createCellStyle(true));
		Assert.assertEquals(style, array.getCellStyle());
		
		//only one item
		arrays = sheet1.getColumnArrayIterator();
		Assert.assertTrue(arrays.hasNext());
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(0, array.getLastIndex());
		Assert.assertFalse(arrays.hasNext());
		
		//add 100, 50, 1, 49, 101
		sheet1.getColumn(100).setWidth(300);
		sheet1.getColumn(50).setWidth(300);
		sheet1.getColumn(1).setWidth(300);
		sheet1.getColumn(49).setWidth(300);
		sheet1.getColumn(101).setWidth(300);
		
		//check the rang , 
		//0, 1, 2-48, 49, 50, 51-99, 100, 101
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(0, array.getLastIndex());
		Assert.assertEquals(150, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(1, array.getIndex());
		Assert.assertEquals(1, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(2, array.getIndex());
		Assert.assertEquals(48, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(49, array.getIndex());
		Assert.assertEquals(49, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(50, array.getIndex());
		Assert.assertEquals(50, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(51, array.getIndex());
		Assert.assertEquals(99, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(100, array.getIndex());
		Assert.assertEquals(100, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(101, array.getIndex());
		Assert.assertEquals(101, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
		Assert.assertFalse(arrays.hasNext());
	}
	
	@Test
	public void testColumnArray2(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Assert.assertNull(sheet1.getColumnArray(9));
		SColumn column = sheet1.getColumn(9);
		Assert.assertNull(sheet1.getColumnArray(9));
		column.setWidth(150);
		SColumnArray array;
		Assert.assertNotNull(array = sheet1.getColumnArray(9));
		Assert.assertEquals(150, array.getWidth());
		
		Assert.assertEquals(false, array.isHidden());
		column.setHidden(true);
		Assert.assertEquals(true, array.isHidden());
		
		SCellStyle style;
		Assert.assertEquals(book.getDefaultCellStyle(), array.getCellStyle());
		column.setCellStyle(style = book.createCellStyle(true));
		Assert.assertEquals(style, array.getCellStyle());
		
		
		//check empty one
		Assert.assertEquals(sheet1.getColumnArray(0),array = sheet1.getColumnArray(8));
		Assert.assertEquals(false, array.isHidden());
		Assert.assertEquals(book.getDefaultCellStyle(), array.getCellStyle());
		
		
		//insert between
		column = sheet1.getColumn(5);
		column.setWidth(160);
		Assert.assertNotNull(array = sheet1.getColumnArray(5));
		Assert.assertEquals(160, array.getWidth());
		
		Assert.assertEquals(false, array.isHidden());
		column.setHidden(true);
		Assert.assertEquals(true, array.isHidden());
		
		Assert.assertEquals(book.getDefaultCellStyle(), array.getCellStyle());
		column.setCellStyle(style = book.createCellStyle(true));
		Assert.assertEquals(style, array.getCellStyle());
		
		
		//check empty one
		Assert.assertNotSame(sheet1.getColumnArray(0),sheet1.getColumnArray(8));
		array = sheet1.getColumnArray(0);
		Assert.assertEquals(false, array.isHidden());
		Assert.assertEquals(book.getDefaultCellStyle(), array.getCellStyle());
		array = sheet1.getColumnArray(8);
		Assert.assertEquals(false, array.isHidden());
		Assert.assertEquals(book.getDefaultCellStyle(), array.getCellStyle());
		
		column = sheet1.getColumn(100);
		column.setWidth(400);
	}
	
	@Test
	public void testInsertColumnArray(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		sheet1.insertColumn(10, 12);
		
		arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		//0-99,100
		sheet1.getColumn(100).setWidth(300);
		sheet1.insertColumn(200, 202);
		
		arrays = sheet1.getColumnArrayIterator();
		SColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(99, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		array.setWidth(44);//change it width to 44, to verify insert split
		Assert.assertEquals(44, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(100, array.getIndex());
		Assert.assertEquals(100, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
		
		
		//0-49,50-52,53-102,103
		sheet1.insertColumn(50, 52);
		
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(49, array.getLastIndex());
		Assert.assertEquals(44, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(50, array.getIndex());
		Assert.assertEquals(52, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());//default width for new insert
		
		array = arrays.next();
		Assert.assertEquals(53, array.getIndex());
		Assert.assertEquals(102, array.getLastIndex());
		Assert.assertEquals(44, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(103, array.getIndex());
		Assert.assertEquals(103, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
		//0-49,50-52,53-102,103 ->
		//0-1,2-51,52-54,55-104,105
		sheet1.insertColumn(0, 1);
		//0-2,3-51,52-54,55-103,104-106,107,108
		sheet1.insertColumn(104, 106);
		
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(1, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(2, array.getIndex());
		Assert.assertEquals(51, array.getLastIndex());
		Assert.assertEquals(44, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(52, array.getIndex());
		Assert.assertEquals(54, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());//default width for new insert
		
		array = arrays.next();
		Assert.assertEquals(55, array.getIndex());
		Assert.assertEquals(103, array.getLastIndex());
		Assert.assertEquals(44, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(104, array.getIndex());
		Assert.assertEquals(106, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(107, array.getIndex());
		Assert.assertEquals(107, array.getLastIndex());
		Assert.assertEquals(44, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(108, array.getIndex());
		Assert.assertEquals(108, array.getLastIndex());
		Assert.assertEquals(300, array.getWidth());
		
	}
	@Test
	public void testColumnIterator(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		Iterator<SColumn> colIter = sheet1.getColumnIterator();
		Assert.assertFalse(colIter.hasNext());
		
		sheet1.getColumn(4).setWidth(20);
		colIter = sheet1.getColumnIterator();
		SColumn col = colIter.next();
		Assert.assertEquals(0, col.getIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		col = colIter.next();
		Assert.assertEquals(1, col.getIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		col = colIter.next();
		Assert.assertEquals(2, col.getIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		col = colIter.next();
		Assert.assertEquals(3, col.getIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		col = colIter.next();
		Assert.assertEquals(4, col.getIndex());
		Assert.assertEquals(20, col.getWidth());
		Assert.assertFalse(colIter.hasNext());
		
		sheet1.getColumn(2).setWidth(30);
		colIter = sheet1.getColumnIterator();
		col = colIter.next();
		Assert.assertEquals(0, col.getIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		col = colIter.next();
		Assert.assertEquals(1, col.getIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		col = colIter.next();
		Assert.assertEquals(2, col.getIndex());
		Assert.assertEquals(30, col.getWidth());
		col = colIter.next();
		Assert.assertEquals(3, col.getIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		col = colIter.next();
		Assert.assertEquals(4, col.getIndex());
		Assert.assertEquals(20, col.getWidth());
		Assert.assertFalse(colIter.hasNext());
	}
	
	@Test
	public void testInsertColumnArray2(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		//0-9,10
		sheet1.getColumn(10).setWidth(33);
		
		//0-9,10,11
		sheet1.insertColumn(10, 10);
		
		arrays = sheet1.getColumnArrayIterator();
		SColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(9, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(10, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(100, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(11, array.getLastIndex());
		Assert.assertEquals(33, array.getWidth());
		
		
		Assert.assertFalse(arrays.hasNext());
	}
	
	@Test
	public void testDeleteColumnArray(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		sheet1.insertColumn(10, 12);
		
		arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		//0-10(111),11-13(222),14(333),15-20(444),21-30(555)
		sheet1.setupColumnArray(0, 10).setWidth(111);
		sheet1.setupColumnArray(11, 13).setWidth(222);
		sheet1.setupColumnArray(14, 14).setWidth(333);
		sheet1.setupColumnArray(15, 20).setWidth(444);
		sheet1.setupColumnArray(21, 30).setWidth(555);
		
		//0-10(111),11-12(222),13-17(444),18-27(555)
		sheet1.deleteColumn(13, 15);
		arrays = sheet1.getColumnArrayIterator();
		SColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(111, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(12, array.getLastIndex());
		Assert.assertEquals(222, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(13, array.getIndex());
		Assert.assertEquals(17, array.getLastIndex());
		Assert.assertEquals(444, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(18, array.getIndex());
		Assert.assertEquals(27, array.getLastIndex());
		Assert.assertEquals(555, array.getWidth());
	}
	
	@Test
	public void testDeleteColumnArray2(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		sheet1.insertColumn(10, 12);
		
		arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		//0-10(111),11-13(222),14(333),15-20(444),21-30(555)
		sheet1.setupColumnArray(0, 10).setWidth(111);
		sheet1.setupColumnArray(11, 13).setWidth(222);
		sheet1.setupColumnArray(14, 14).setWidth(333);
		sheet1.setupColumnArray(15, 20).setWidth(444);
		sheet1.setupColumnArray(21, 30).setWidth(555);
		
		//0-10(111),11-11(222),12-15(444),16-25(555)
		sheet1.deleteColumn(12, 16);
		arrays = sheet1.getColumnArrayIterator();
		SColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(111, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(11, array.getLastIndex());
		Assert.assertEquals(222, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(12, array.getIndex());
		Assert.assertEquals(15, array.getLastIndex());
		Assert.assertEquals(444, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(16, array.getIndex());
		Assert.assertEquals(25, array.getLastIndex());
		Assert.assertEquals(555, array.getWidth());
		Assert.assertFalse(arrays.hasNext());
	}
	
	@Test
	public void testDeleteColumnArray3(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Iterator<SColumnArray> arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		sheet1.insertColumn(10, 12);
		
		arrays = sheet1.getColumnArrayIterator();
		Assert.assertFalse(arrays.hasNext());
		
		//0-10(111),11-20(222)
		sheet1.setupColumnArray(0, 10).setWidth(111);
		sheet1.setupColumnArray(11, 20).setWidth(222);
		
		
		//0-10(111),11-18(222)
		sheet1.deleteColumn(12,13);
		arrays = sheet1.getColumnArrayIterator();
		SColumnArray array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(111, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(18, array.getLastIndex());
		Assert.assertEquals(222, array.getWidth());
		Assert.assertFalse(arrays.hasNext());
		
		
		//0-10(111),11-15(222)
		sheet1.deleteColumn(11, 13);
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(10, array.getLastIndex());
		Assert.assertEquals(111, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(11, array.getIndex());
		Assert.assertEquals(15, array.getLastIndex());
		Assert.assertEquals(222, array.getWidth());	
		Assert.assertFalse(arrays.hasNext());
		
		//0-9(111),10-12(222)
		sheet1.deleteColumn(10, 12);
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(9, array.getLastIndex());
		Assert.assertEquals(111, array.getWidth());
		
		array = arrays.next();
		Assert.assertEquals(10, array.getIndex());
		Assert.assertEquals(12, array.getLastIndex());
		Assert.assertEquals(222, array.getWidth());
		
		Assert.assertFalse(arrays.hasNext());
		
		//0-9(111),10-12(222)
		sheet1.deleteColumn(10, 12);
		arrays = sheet1.getColumnArrayIterator();
		array = arrays.next();
		Assert.assertEquals(0, array.getIndex());
		Assert.assertEquals(9, array.getLastIndex());
		Assert.assertEquals(111, array.getWidth());
		Assert.assertFalse(arrays.hasNext());		
		
	}
	
	@Test
	public void testRowColumn(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.setDefaultColumnWidth(100);
		sheet1.setDefaultRowHeight(200);
		Assert.assertEquals(100, sheet1.getDefaultColumnWidth());
		Assert.assertEquals(200, sheet1.getDefaultRowHeight());
		
		Assert.assertEquals(100, sheet1.getColumn(0).getWidth());
		Assert.assertEquals(200, sheet1.getRow(1).getHeight());
		
		sheet1.getColumn(0).setWidth(30);
		sheet1.getRow(1).setHeight(60);
		
		Assert.assertEquals(100, sheet1.getColumn(2).getWidth());
		Assert.assertEquals(200, sheet1.getRow(2).getHeight());
		Assert.assertEquals(false, sheet1.getColumn(2).isHidden());
		Assert.assertEquals(false, sheet1.getRow(2).isHidden());
		Assert.assertEquals(30, sheet1.getColumn(0).getWidth());
		Assert.assertEquals(60, sheet1.getRow(1).getHeight());
		
		sheet1.getColumn(100).setHidden(true); //mark 2nd column
		sheet1.getRow(1000).setHidden(true); //mark 2nd row
		
		sheet1.getCell(600, 60).setValue("Test");//this will not create custom column
		
		
		Assert.assertEquals(true, sheet1.getColumn(100).isHidden());
		Assert.assertEquals(true, sheet1.getRow(1000).isHidden());
		
		Iterator<SRow> rowiter = sheet1.getRowIterator();
		Assert.assertTrue(rowiter.hasNext());
		SRow row = rowiter.next();
		Assert.assertEquals(1, row.getIndex());
		Assert.assertEquals(60, row.getHeight());
		Assert.assertEquals(false, row.isHidden());
		
		Assert.assertTrue(rowiter.hasNext());
		row = rowiter.next();
		Assert.assertEquals(600, row.getIndex());
		Assert.assertEquals(sheet1.getDefaultRowHeight(), row.getHeight());
		Assert.assertEquals(false, row.isHidden());
		
		Assert.assertTrue(rowiter.hasNext());
		row = rowiter.next();
		Assert.assertEquals(1000, row.getIndex());
		Assert.assertEquals(sheet1.getDefaultRowHeight(), row.getHeight());
		Assert.assertEquals(true, row.isHidden());
		
		Assert.assertFalse(rowiter.hasNext());
		
		
		Iterator<SColumnArray> coliter = sheet1.getColumnArrayIterator();
		Assert.assertTrue(coliter.hasNext());
		SColumnArray col = coliter.next();
		Assert.assertEquals(0, col.getIndex());
		Assert.assertEquals(0, col.getLastIndex());
		Assert.assertEquals(30, col.getWidth());
		Assert.assertEquals(false, col.isHidden());
		
		Assert.assertTrue(coliter.hasNext());
		col = coliter.next();
		Assert.assertEquals(1, col.getIndex());
		Assert.assertEquals(99, col.getLastIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		Assert.assertEquals(false, col.isHidden());
		
		Assert.assertTrue(coliter.hasNext());
		col = coliter.next();
		Assert.assertEquals(100, col.getIndex());
		Assert.assertEquals(100, col.getLastIndex());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), col.getWidth());
		Assert.assertEquals(true, col.isHidden());
		
		Assert.assertFalse(coliter.hasNext());
		
	}
	
	
	static String asString(SRow row){
		return Integer.toString(row.getIndex()+1);
	}
	static String asString(SColumn column){
		return CellReference.convertNumToColString(column.getIndex());
	}
	
	@Test
	public void testReferenceString(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		Assert.assertEquals("1",asString(sheet1.getRow(0)));
		Assert.assertEquals("101",asString(sheet1.getRow(100)));
		Assert.assertEquals("A",asString(sheet1.getColumn(0)));
		Assert.assertEquals("AY",asString(sheet1.getColumn(50)));
		Assert.assertEquals("A1",sheet1.getCell(0,0).getReferenceString());
		Assert.assertEquals("AY101",sheet1.getCell(100,50).getReferenceString());
//		Assert.assertEquals("Sheet1!A1",sheet1.getCell(0,0).getReferenceString(true));
//		Assert.assertEquals("Sheet1!AY101",sheet1.getCell(100,50).getReferenceString(true));
		
		
		sheet1.getCell(9, 5).setValue("(9,5)");
		
		Assert.assertEquals("10",asString(sheet1.getRow(9)));
		Assert.assertEquals("F",asString(sheet1.getColumn(5)));
		Assert.assertEquals("F10",sheet1.getCell(9,5).getReferenceString());
//		Assert.assertEquals("Sheet1!F10",sheet1.getCell(9,5).getReferenceString(true));
		
//		dump(book);
	}
	
	@Test
	public void testCellRange(){
		SBook book = SBooks.createBook("book1");
		initialDataGrid(book.createSheet("Sheet1"));
		Assert.assertEquals(1, book.getNumOfSheet());
		SSheet sheet = initialDataGrid(book.createSheet("Sheet2"));
		Assert.assertEquals(-1, sheet.getStartRowIndex());
		Assert.assertEquals(-1, sheet.getEndRowIndex());
		Assert.assertEquals(-1, sheet.getStartColumnIndex());
		Assert.assertEquals(-1, sheet.getEndColumnIndex());
		Assert.assertEquals(-1, sheet.getStartCellIndex(0));
		Assert.assertEquals(-1, sheet.getEndCellIndex(0));
		
		SRow row = sheet.getRow(3);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(-1, sheet.getStartCellIndex(row.getIndex()));
		Assert.assertEquals(-1, sheet.getEndCellIndex(row.getIndex()));
		SColumn column = sheet.getColumn(6);
		Assert.assertEquals(true, column.isNull());
		
		SCell cell = sheet.getCell(3, 6);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(3,6)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(true, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(3,6)", cell.getValue());
		
		column.setWidth(300);
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(3, sheet.getEndRowIndex());
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(6, sheet.getEndColumnIndex());
		
		Assert.assertEquals(-1, sheet.getStartCellIndex(0));
		Assert.assertEquals(-1, sheet.getEndCellIndex(0));
		Assert.assertEquals(6, sheet.getStartCellIndex(3));
		Assert.assertEquals(6, sheet.getEndCellIndex(3));
		Assert.assertEquals(6, sheet.getStartCellIndex(row.getIndex()));
		Assert.assertEquals(6, sheet.getEndCellIndex(row.getIndex()));
		Assert.assertEquals(-1, sheet.getStartCellIndex(4));
		Assert.assertEquals(-1, sheet.getEndCellIndex(4));
		
		
		//another cell
		column = sheet.getColumn(12);
		Assert.assertEquals(true, column.isNull());
		
		cell = sheet.getCell(3, 12);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(3,12)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(true, column.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(3,12)", cell.getValue());
		
		column.setWidth(300);
		Assert.assertEquals(false, column.isNull());
		
		Assert.assertEquals(6, sheet.getStartCellIndex(row.getIndex()));
		Assert.assertEquals(12, sheet.getEndCellIndex(row.getIndex()));
		
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(3, sheet.getEndRowIndex());
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(12, sheet.getEndColumnIndex());
		Assert.assertEquals(-1, sheet.getStartCellIndex(0));
		Assert.assertEquals(-1, sheet.getEndCellIndex(0));
		Assert.assertEquals(6, sheet.getStartCellIndex(3));
		Assert.assertEquals(12, sheet.getEndCellIndex(3));
		Assert.assertEquals(-1, sheet.getStartCellIndex(4));
		Assert.assertEquals(-1, sheet.getEndCellIndex(4));
		
		
		//another cell
		row = sheet.getRow(4);
		column = sheet.getColumn(8);
		Assert.assertEquals(true, row.isNull());
		Assert.assertEquals(false, column.isNull());
		Assert.assertEquals(true, sheet.getColumn(13).isNull());
		
		cell = sheet.getCell(4, 8);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(4,8)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(4,8)", cell.getValue());
		
		column.setWidth(300);
		Assert.assertEquals(false, column.isNull());
		
		Assert.assertEquals(8, sheet.getStartCellIndex(row.getIndex()));
		Assert.assertEquals(8, sheet.getEndCellIndex(row.getIndex()));
		
		Assert.assertEquals(3, sheet.getStartRowIndex());
		Assert.assertEquals(4, sheet.getEndRowIndex());
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(12, sheet.getEndColumnIndex());
		Assert.assertEquals(-1, sheet.getStartCellIndex(0));
		Assert.assertEquals(-1, sheet.getEndCellIndex(0));
		Assert.assertEquals(6, sheet.getStartCellIndex(3));
		Assert.assertEquals(12, sheet.getEndCellIndex(3));
		Assert.assertEquals(8, sheet.getStartCellIndex(4));
		Assert.assertEquals(8, sheet.getEndCellIndex(4));
		
		
		//another cell
		row = sheet.getRow(0);
		column = sheet.getColumn(0);
		Assert.assertEquals(true, row.isNull());
		
		cell = sheet.getCell(0, 0);
		Assert.assertEquals(true, cell.isNull());
		
		cell.setValue("(0,0)");
		Assert.assertEquals(false, row.isNull());
		Assert.assertEquals(false, cell.isNull());
		Assert.assertEquals("(0,0)", cell.getValue());
		
		column.setWidth(300);
		Assert.assertEquals(false, column.isNull());
		
		Assert.assertEquals(0, sheet.getStartCellIndex(row.getIndex()));
		Assert.assertEquals(0, sheet.getEndCellIndex(row.getIndex()));
		
		Assert.assertEquals(0, sheet.getStartRowIndex());
		Assert.assertEquals(4, sheet.getEndRowIndex());
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(12, sheet.getEndColumnIndex());
		Assert.assertEquals(0, sheet.getStartCellIndex(0));
		Assert.assertEquals(0, sheet.getEndCellIndex(0));
		Assert.assertEquals(6, sheet.getStartCellIndex(3));
		Assert.assertEquals(12, sheet.getEndCellIndex(3));
		Assert.assertEquals(8, sheet.getStartCellIndex(4));
		Assert.assertEquals(8, sheet.getEndCellIndex(4));	
	}
	
	@Test
	public void testClearSheetCell(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		
		for(int i=10;i<=20;i+=2){
			for(int j=3;j<=15;j+=3){
				SCell cell = sheet.getCell(i, j);
				cell.setValue("("+i+","+j+")");
				sheet.getColumn(j).setWidth(300);
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
		
//		dump(book);
		
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
		
		Assert.assertEquals(0, sheet.getStartColumnIndex());
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
		
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(15, sheet.getEndColumnIndex());

		
	}
	
	@Test
	public void testInsertDeleteRow(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		
		for(int i=10;i<=20;i+=2){
			for(int j=3;j<=15;j+=3){
				SCell cell = sheet.getCell(i, j);
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
		
		Assert.assertEquals(10, sheet.getStartRowIndex());
		Assert.assertEquals(20, sheet.getEndRowIndex());
		
		SRow row10 = sheet.getRow(10);
		SRow row12 = sheet.getRow(12);
		SRow row14 = sheet.getRow(14);
		SRow row16 = sheet.getRow(16);
		
		sheet.insertRow(12, 14);
		
		Assert.assertEquals(false, sheet.getRow(10).isNull());
		Assert.assertEquals(true, sheet.getRow(12).isNull());
		Assert.assertEquals(true, sheet.getRow(14).isNull());
		Assert.assertEquals(true, sheet.getRow(16).isNull());
		
		Assert.assertEquals(10, row10.getIndex());
		Assert.assertEquals(15, row12.getIndex());
		Assert.assertEquals(17, row14.getIndex());
		Assert.assertEquals(19, row16.getIndex());
		
		
		Assert.assertEquals(row10, sheet.getRow(10));
		Assert.assertEquals(row12, sheet.getRow(15));
		Assert.assertEquals(row14, sheet.getRow(17));
		Assert.assertEquals(row16, sheet.getRow(19));
		
		Assert.assertEquals("(10,3)", sheet.getCell(10, 3).getValue());
		Assert.assertEquals("(12,6)", sheet.getCell(15, 6).getValue());
		Assert.assertEquals("(14,9)", sheet.getCell(17, 9).getValue());
		Assert.assertEquals("(16,12)", sheet.getCell(19, 12).getValue());
		
		Assert.assertEquals(10, sheet.getStartRowIndex());
		Assert.assertEquals(23, sheet.getEndRowIndex());
		
		sheet.insertRow(100, 102);
		
		Assert.assertEquals(10, sheet.getStartRowIndex());
		Assert.assertEquals(23, sheet.getEndRowIndex());
		
		
		sheet.deleteRow(10, 15);
		
		Assert.assertEquals(true, sheet.getRow(10).isNull());
		Assert.assertEquals(true, sheet.getRow(12).isNull());
		Assert.assertEquals(true, sheet.getRow(14).isNull());
		Assert.assertEquals(true, sheet.getRow(16).isNull());
		
		try{
			row10.getIndex();
			Assert.fail("orphan");
		}catch(IllegalStateException ex){}
		try{
			row12.getIndex();
		}catch(IllegalStateException ex){}
		Assert.assertEquals(11, row14.getIndex());
		Assert.assertEquals(13, row16.getIndex());
		
		
		Assert.assertEquals(row14, sheet.getRow(11));
		Assert.assertEquals(row16, sheet.getRow(13));
		
		Assert.assertEquals(null, sheet.getCell(10, 3).getValue());
		Assert.assertEquals(null, sheet.getCell(12, 6).getValue());
		Assert.assertEquals("(14,9)", sheet.getCell(11, 9).getValue());
		Assert.assertEquals("(16,12)", sheet.getCell(13, 12).getValue());
		
		Assert.assertEquals(11, sheet.getStartRowIndex());
		Assert.assertEquals(17, sheet.getEndRowIndex());
		
		
		sheet.deleteRow(100, 102);
		
		Assert.assertEquals(11, sheet.getStartRowIndex());
		Assert.assertEquals(17, sheet.getEndRowIndex());
	}
	
	@Test
	public void testInsertDeleteColumn(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		
		for(int i=10;i<=20;i+=2){
			for(int j=3;j<=15;j+=3){
				SCell cell = sheet.getCell(j, i);
				cell.setValue("("+j+","+i+")");//will not create custom column
			}
		}
		Assert.assertEquals(true, sheet.getColumn(10).isNull());
		Assert.assertEquals(true, sheet.getColumn(12).isNull());
		Assert.assertEquals(true, sheet.getColumn(14).isNull());
		Assert.assertEquals(true, sheet.getColumn(16).isNull());
		Assert.assertEquals(true, sheet.getColumn(20).isNull());
		
		sheet.getColumn(10).setWidth(300);
		sheet.getColumn(12).setWidth(300);
		sheet.getColumn(14).setWidth(300);
		sheet.getColumn(16).setWidth(300);
		sheet.getColumn(20).setWidth(300);
		
		Assert.assertEquals(false, sheet.getColumn(10).isNull());
		Assert.assertEquals(false, sheet.getColumn(12).isNull());
		Assert.assertEquals(false, sheet.getColumn(14).isNull());
		Assert.assertEquals(false, sheet.getColumn(16).isNull());
		Assert.assertEquals(false, sheet.getColumn(20).isNull());
		
		Assert.assertEquals("(3,10)", sheet.getCell(3, 10).getValue());
		Assert.assertEquals("(6,12)", sheet.getCell(6, 12).getValue());
		Assert.assertEquals("(9,14)", sheet.getCell(9, 14).getValue());
		Assert.assertEquals("(12,16)", sheet.getCell(12, 16).getValue());
		
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(20, sheet.getEndColumnIndex());
		
		SColumn column10 = sheet.getColumn(10);
		SColumn column12 = sheet.getColumn(12);
		SColumn column14 = sheet.getColumn(14);
		SColumn column16 = sheet.getColumn(16);
		
		sheet.insertColumn(12, 14);
		
		Assert.assertEquals(false, sheet.getColumn(10).isNull());
		Assert.assertEquals(false, sheet.getColumn(12).isNull());
		Assert.assertEquals(false, sheet.getColumn(14).isNull());
		Assert.assertEquals(false, sheet.getColumn(16).isNull());
		
		//no more avaiable test after re design to column array
//		Assert.assertEquals(10, column10.getIndex());
//		Assert.assertEquals(15, column12.getIndex());
//		Assert.assertEquals(17, column14.getIndex());
//		Assert.assertEquals(19, column16.getIndex());
//		
//		
//		Assert.assertEquals(column10, sheet.getColumn(10));
//		Assert.assertEquals(column12, sheet.getColumn(15));
//		Assert.assertEquals(column14, sheet.getColumn(17));
//		Assert.assertEquals(column16, sheet.getColumn(19));
		
		Assert.assertEquals("(3,10)", sheet.getCell(3,10).getValue());
		Assert.assertEquals("(6,12)", sheet.getCell(6,15).getValue());
		Assert.assertEquals("(9,14)", sheet.getCell(9,17).getValue());
		Assert.assertEquals("(12,16)", sheet.getCell(12,19).getValue());
		
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(23, sheet.getEndColumnIndex());
		
		sheet.insertColumn(100, 102);
		
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(23, sheet.getEndColumnIndex());
		
		
		sheet.deleteColumn(10, 15);
		
		//no more avaiable test after re design to column array
//		Assert.assertEquals(true, sheet.getColumn(10).isNull());
//		Assert.assertEquals(true, sheet.getColumn(12).isNull());
//		Assert.assertEquals(true, sheet.getColumn(14).isNull());
//		Assert.assertEquals(true, sheet.getColumn(16).isNull());
		
//		try{
//			column10.getIndex();
//		}catch(IllegalStateException ex){}
//		try{
//			column12.getIndex();
//		}catch(IllegalStateException ex){}
		
		//no more avaiable test after re design to column array
//		Assert.assertEquals(11, column14.getIndex());
//		Assert.assertEquals(13, column16.getIndex());
		
//		Assert.assertEquals(column14, sheet.getColumn(11));
//		Assert.assertEquals(column16, sheet.getColumn(13));
		
		Assert.assertEquals(null, sheet.getCell(3,10).getValue());
		Assert.assertEquals(null, sheet.getCell(6,12).getValue());
		Assert.assertEquals("(9,14)", sheet.getCell(9,11).getValue());
		Assert.assertEquals("(12,16)", sheet.getCell(12,13).getValue());
		
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(17, sheet.getEndColumnIndex());
		
		
		sheet.deleteColumn(100, 102);
		
		Assert.assertEquals(0, sheet.getStartColumnIndex());
		Assert.assertEquals(17, sheet.getEndColumnIndex());
	}
	
	public static void dump(SBook book){
		StringBuilder builder = new StringBuilder();
		((BookImpl)book).dump(builder);
		System.out.println(builder.toString());
	}
	
	@Test
	public void testStyle(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		
		SCellStyle style = book.getDefaultCellStyle();
		
		Assert.assertEquals(style, sheet.getRow(10).getCellStyle());
		Assert.assertEquals(style, sheet.getColumn(3).getCellStyle());
		Assert.assertEquals(style, sheet.getCell(10,3).getCellStyle());
		
		
		SCellStyle cellStyle = book.createCellStyle(true);
		sheet.getCell(10, 3).setCellStyle(cellStyle);
		
		Assert.assertEquals(style, sheet.getRow(10).getCellStyle());
		Assert.assertEquals(style, sheet.getColumn(3).getCellStyle());
		Assert.assertEquals(cellStyle, sheet.getCell(10,3).getCellStyle());
		
		SCellStyle rowStyle = book.createCellStyle(true);
		sheet.getRow(9).setCellStyle(rowStyle);
		
		Assert.assertEquals(style, sheet.getRow(10).getCellStyle());
		Assert.assertEquals(style, sheet.getColumn(3).getCellStyle());
		Assert.assertEquals(cellStyle, sheet.getCell(10,3).getCellStyle());
		
		Assert.assertEquals(rowStyle, sheet.getRow(9).getCellStyle());
		Assert.assertEquals(rowStyle, sheet.getCell(9,3).getCellStyle());
		
		
		SCellStyle columnStyle = book.createCellStyle(true);
		sheet.getColumn(4).setCellStyle(columnStyle);
		
		Assert.assertEquals(style, sheet.getRow(10).getCellStyle());
		Assert.assertEquals(style, sheet.getColumn(3).getCellStyle());
		Assert.assertEquals(cellStyle, sheet.getCell(10,3).getCellStyle());
		
		Assert.assertEquals(rowStyle, sheet.getRow(9).getCellStyle());
		Assert.assertEquals(rowStyle, sheet.getCell(9,3).getCellStyle());
		
		Assert.assertEquals(columnStyle, sheet.getColumn(4).getCellStyle());
		Assert.assertEquals(rowStyle, sheet.getCell(9,4).getCellStyle());//style on row 9 first.
		Assert.assertEquals(columnStyle, sheet.getCell(10,4).getCellStyle());
		
	}
	
	
	@Test
	public void testStyleSearch(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		
		SCellStyle style1 = book.createCellStyle(true);
		CellStyleMatcher matcher = new CellStyleMatcher(book.createCellStyle(false));//a style not in table
		
		Assert.assertEquals(book.getDefaultCellStyle(),book.searchCellStyle(matcher));
		
		Assert.assertNotSame(style1, book.getDefaultCellStyle());
		
		
		
		style1.setAlignment(Alignment.CENTER);
		matcher.setAlignment(Alignment.CENTER);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setAlignment(Alignment.RIGHT);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setAlignment(Alignment.RIGHT);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		
		style1.setFillColor(book.createColor("#FF0F0B"));
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFillColor("#FF0F0B");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		//////////
		style1.setBorderBottom(BorderType.MEDIUM);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setBorderBottom(BorderType.MEDIUM);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setBorderBottomColor(book.createColor("#FF0000"));
		matcher.setBorderBottomColor("#FF00FF");//will hit if didn't set color, because at the begin the border-type is none - that cause matcher ignore the color mapping
		Assert.assertNull(book.searchCellStyle(matcher)); // 
		matcher.setBorderBottomColor("#FF0000");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setBorderLeft(BorderType.MEDIUM);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setBorderLeft(BorderType.MEDIUM);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setBorderLeftColor(book.createColor("#FF0000"));
		matcher.setBorderLeftColor("#FF00FF");
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setBorderLeftColor("#FF0000");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setBorderRight(BorderType.MEDIUM);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setBorderRight(BorderType.MEDIUM);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setBorderRightColor(book.createColor("#FF0000"));
		matcher.setBorderRightColor("#FF00FF");
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setBorderRightColor("#FF0000");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setBorderTop(BorderType.MEDIUM);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setBorderTop(BorderType.MEDIUM);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setBorderTopColor(book.createColor("#FF0000"));
		matcher.setBorderTopColor("#FF00FF");
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setBorderTopColor("#FF0000");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setDataFormat("yyyymd");
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setDataFormat("yyyymd");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setFillPattern(FillPattern.SOLID_FOREGROUND);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFillPattern(FillPattern.SOLID_FOREGROUND);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setBoldweight(Boldweight.BOLD);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontBoldweight(Boldweight.BOLD);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setColor(book.createColor("#0000FF"));
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontColor("#0000FF");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setHeightPoints(26);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontHeightPoints(26);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setItalic(true);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontItalic(true);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setName("system");
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontName("system");
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setStrikeout(true);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontStrikeout(true);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setTypeOffset(TypeOffset.SUB);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontTypeOffset(TypeOffset.SUB);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.getFont().setUnderline(Underline.SINGLE);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setFontUnderline(Underline.SINGLE);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setHidden(true);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setHidden(true);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setLocked(false);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setLocked(false);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
		style1.setWrapText(true);
		Assert.assertNull(book.searchCellStyle(matcher));
		matcher.setWrapText(true);
		Assert.assertEquals(style1,book.searchCellStyle(matcher));
		
	}
	
	@Test
	public void testFontSearch(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		
		SFont font1 = book.createFont(true);
		FontMatcher matcher = new FontMatcher(book.createFont(false));//a style not in table
		
		Assert.assertEquals(book.getDefaultFont(),book.searchFont(matcher));
		
		Assert.assertNotSame(font1, book.getDefaultFont());
		

		font1.setBoldweight(Boldweight.BOLD);
		matcher.setBoldweight(Boldweight.BOLD);
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		font1.setColor(book.createColor("#0000FF"));
		Assert.assertNull(book.searchFont(matcher));
		matcher.setColor("#0000FF");
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		font1.setHeightPoints(26);
		Assert.assertNull(book.searchFont(matcher));
		matcher.setHeightPoints(26);
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		font1.setItalic(true);
		Assert.assertNull(book.searchFont(matcher));
		matcher.setItalic(true);
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		font1.setName("system");
		Assert.assertNull(book.searchFont(matcher));
		matcher.setName("system");
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		font1.setStrikeout(true);
		Assert.assertNull(book.searchFont(matcher));
		matcher.setStrikeout(true);
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		font1.setTypeOffset(TypeOffset.SUB);
		Assert.assertNull(book.searchFont(matcher));
		matcher.setTypeOffset(TypeOffset.SUB);
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		font1.setUnderline(Underline.SINGLE);
		Assert.assertNull(book.searchFont(matcher));
		matcher.setUnderline(Underline.SINGLE);
		Assert.assertEquals(font1,book.searchFont(matcher));
		
		
		
	}
	
	@Test
	public void testGeneralCellValue(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		Date now = new Date();
		ErrorValue err = new ErrorValue(ErrorValue.INVALID_FORMULA);
		SCell cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		cell.setValue("abc");
		Assert.assertFalse(cell.isFormulaParsingError());
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("abc",cell.getValue());
		
		cell.setValue(123);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(123D,cell.getValue());
		
		cell.setValue(now);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(now,cell.getDateValue());
		
		cell.setValue(true);
		Assert.assertEquals(CellType.BOOLEAN, cell.getType());
		Assert.assertEquals(true,cell.getValue());
		
		cell.setValue(err);
		Assert.assertEquals(CellType.ERROR, cell.getType());
		Assert.assertEquals(err,cell.getValue());
		
		cell.setValue("=SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999D, cell.getValue());
		
		try{
			cell.setValue("=)))(999)");
			Assert.fail("not here");
		}catch(InvalidModelOpException x){
			Assert.assertEquals(CellType.FORMULA, cell.getType());
			Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
			Assert.assertEquals("SUM(999)", cell.getFormulaValue());
			Assert.assertEquals(999D, cell.getValue());
		}
		
		cell.clearValue();
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		
		//on non cached cell
		cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		cell.setValue("abc");
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("abc",cell.getValue());
		
		cell.setValue(123);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(123D,cell.getValue());
		
		cell.setValue(now);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(now,cell.getDateValue());
		
		cell.setValue(false);
		Assert.assertEquals(CellType.BOOLEAN, cell.getType());
		Assert.assertEquals(false,cell.getValue());
		
		
		cell.setValue(err);
		Assert.assertEquals(CellType.ERROR, cell.getType());
		Assert.assertEquals(err,cell.getValue());
		
		cell.setValue("=SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999D, cell.getValue());
		
		cell.clearValue();
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
	}
	
	@Test
	public void testGeneralCellValue2(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		Date now = new Date();
		ErrorValue err = new ErrorValue(ErrorValue.INVALID_FORMULA);
		SCell cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		cell.setStringValue("abc");
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("abc",cell.getStringValue());
		
		cell.setNumberValue(123D);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(123D,cell.getNumberValue());
		
		cell.setDateValue(now);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(now,cell.getDateValue());
		
		cell.setBooleanValue(true);
		Assert.assertEquals(CellType.BOOLEAN, cell.getType());
		Assert.assertEquals(Boolean.TRUE,cell.getBooleanValue());
		
		cell.setErrorValue(err);
		Assert.assertEquals(CellType.ERROR, cell.getType());
		Assert.assertEquals(err,cell.getErrorValue());
		
		cell.setFormulaValue("SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999D, cell.getNumberValue());
		
		cell.clearValue();
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
	}
	
	@Test
	public void testGeneralCellValueError(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		Date now = new Date();
		ErrorValue err = new ErrorValue(ErrorValue.INVALID_FORMULA);
		SCell cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		Assert.assertEquals("",cell.getStringValue());
		Assert.assertEquals(0.0,cell.getNumberValue().doubleValue());
		cell.getDateValue();
		try{
			cell.getErrorValue();
			Assert.fail();
		}catch(IllegalStateException x){}
		try{
			cell.getFormulaValue();
			Assert.fail();
		}catch(IllegalStateException x){}
		
		
		cell.setFormulaValue("SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999D, cell.getNumberValue());
		
		try{
			cell.getStringValue();
			Assert.fail();
		}catch(IllegalStateException x){}
		try{
			cell.getErrorValue();
			Assert.fail();
		}catch(IllegalStateException x){}
		
		try{
			cell.setFormulaValue("[(999)");
			Assert.fail("not here");
		}catch(InvalidModelOpException x){
			Assert.assertEquals(CellType.FORMULA, cell.getType());
			Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
			Assert.assertEquals("SUM(999)", cell.getFormulaValue());
			Assert.assertEquals(999D, cell.getNumberValue());
		}
		
		try{
			cell.getStringValue();
			Assert.fail();
		}catch(IllegalStateException x){}
		
		cell.clearValue();
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		Assert.assertEquals("",cell.getStringValue());
		Assert.assertEquals(0.0,cell.getNumberValue().doubleValue());
		cell.getDateValue();
		try{
			cell.getErrorValue();
			Assert.fail();
		}catch(IllegalStateException x){}
		try{
			cell.getFormulaValue();
			Assert.fail();
		}catch(IllegalStateException x){}
	}
	
	@Test
	public void testPicture(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));

		SPicture p1 = sheet.addPicture(Format.PNG, new byte[0], new ViewAnchor(6, 10, 22, 33, 800, 600));
		SPicture p2 = sheet.addPicture(Format.PNG, new byte[0], new ViewAnchor(12, 14, 22, 33, 800, 600));
		
		Assert.assertEquals(2, sheet.getPictures().size());
		Assert.assertEquals(p1,sheet.getPictures().get(0));
		Assert.assertEquals(p2,sheet.getPictures().get(1));
		
		sheet.insertRow(7, 8);
		Assert.assertEquals(6,p1.getAnchor().getRowIndex());
		Assert.assertEquals(14,p2.getAnchor().getRowIndex());
		
		sheet.insertRow(6, 8);
		Assert.assertEquals(9,p1.getAnchor().getRowIndex());
		Assert.assertEquals(17,p2.getAnchor().getRowIndex());
		
		sheet.deleteRow(10, 11);
		Assert.assertEquals(9,p1.getAnchor().getRowIndex());
		Assert.assertEquals(15,p2.getAnchor().getRowIndex());
		Assert.assertEquals(33,p2.getAnchor().getYOffset());
		
		sheet.deleteRow(10, 15);
		Assert.assertEquals(10,p2.getAnchor().getRowIndex());
		Assert.assertEquals(0,p2.getAnchor().getYOffset());
		
		sheet.insertColumn(12, 14);
		Assert.assertEquals(10,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(17,p2.getAnchor().getColumnIndex());
		
		sheet.insertColumn(10, 10);
		Assert.assertEquals(11,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(18,p2.getAnchor().getColumnIndex());
		
		
		sheet.deleteColumn(15, 15);
		Assert.assertEquals(11,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(17,p2.getAnchor().getColumnIndex());
		
		sheet.deleteColumn(15, 16);
		Assert.assertEquals(11,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(15,p2.getAnchor().getColumnIndex());
		Assert.assertEquals(22,p2.getAnchor().getXOffset());
		
		sheet.deleteColumn(10, 19);
		Assert.assertEquals(10,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(0,p1.getAnchor().getXOffset());
		Assert.assertEquals(10,p2.getAnchor().getColumnIndex());
		Assert.assertEquals(0,p2.getAnchor().getXOffset());
		
	}
	
	@Test
	public void testChart(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));

		//no chart data implement yet
		SChart p1 = sheet.addChart(SChart.ChartType.BAR, new ViewAnchor(6, 10, 22, 33, 800, 600));
		SChart p2 = sheet.addChart(SChart.ChartType.BAR, new ViewAnchor(12, 14, 22, 33, 800, 600));
		
		p1.setTitle("MyChart");
		p1.setXAxisTitle("X");
		p1.setYAxisTitle("Y");
		
		Assert.assertEquals("MyChart", p1.getTitle());
		Assert.assertEquals("X", p1.getXAxisTitle());
		Assert.assertEquals("Y", p1.getYAxisTitle());
		
		Assert.assertEquals(6, p1.getAnchor().getRowIndex());
		Assert.assertEquals(10, p1.getAnchor().getColumnIndex());
		Assert.assertEquals(22, p1.getAnchor().getXOffset());
		Assert.assertEquals(33, p1.getAnchor().getYOffset());
		Assert.assertEquals(800, p1.getAnchor().getWidth());
		Assert.assertEquals(600, p1.getAnchor().getHeight());
		
		
		Assert.assertEquals(2, sheet.getCharts().size());
		Assert.assertEquals(p1,sheet.getCharts().get(0));
		Assert.assertEquals(p2,sheet.getCharts().get(1));
		
		sheet.insertRow(7, 8);
		Assert.assertEquals(6,p1.getAnchor().getRowIndex());
		Assert.assertEquals(14,p2.getAnchor().getRowIndex());
		
		sheet.insertRow(6, 8);
		Assert.assertEquals(9,p1.getAnchor().getRowIndex());
		Assert.assertEquals(17,p2.getAnchor().getRowIndex());
		
		sheet.deleteRow(10, 11);
		Assert.assertEquals(9,p1.getAnchor().getRowIndex());
		Assert.assertEquals(15,p2.getAnchor().getRowIndex());
		Assert.assertEquals(33,p2.getAnchor().getYOffset());
		
		sheet.deleteRow(10, 15);
		Assert.assertEquals(10,p2.getAnchor().getRowIndex());
		Assert.assertEquals(0,p2.getAnchor().getYOffset());
		
		sheet.insertColumn(12, 14);
		Assert.assertEquals(10,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(17,p2.getAnchor().getColumnIndex());
		
		sheet.insertColumn(10, 10);
		Assert.assertEquals(11,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(18,p2.getAnchor().getColumnIndex());
		
		
		sheet.deleteColumn(15, 15);
		Assert.assertEquals(11,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(17,p2.getAnchor().getColumnIndex());
		
		sheet.deleteColumn(15, 16);
		Assert.assertEquals(11,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(15,p2.getAnchor().getColumnIndex());
		Assert.assertEquals(22,p2.getAnchor().getXOffset());
		
		sheet.deleteColumn(10, 19);
		Assert.assertEquals(10,p1.getAnchor().getColumnIndex());
		Assert.assertEquals(0,p1.getAnchor().getXOffset());
		Assert.assertEquals(10,p2.getAnchor().getColumnIndex());
		Assert.assertEquals(0,p2.getAnchor().getXOffset());
		
	}
	
	@Test
	public void testChartData(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		sheet.getCell(0, 0).setStringValue("A");
		sheet.getCell(1, 0).setStringValue("B");
		sheet.getCell(2, 0).setStringValue("C");
		sheet.getCell(0, 3).setStringValue("My Series");
		sheet.getCell(0, 1).setNumberValue(1.0);
		sheet.getCell(1, 1).setNumberValue(2.0);
		sheet.getCell(2, 1).setNumberValue(3.0);
		sheet.getCell(0, 2).setNumberValue(4.0);
		sheet.getCell(1, 2).setNumberValue(5.0);
		sheet.getCell(2, 2).setNumberValue(6.0);

		
		SChart p1 = sheet.addChart(SChart.ChartType.BAR, new ViewAnchor(6, 10, 22, 33, 800, 600));
		
		SGeneralChartData chartData = (SGeneralChartData)p1.getData();
		Assert.assertEquals(0, chartData.getNumOfCategory());
		Assert.assertEquals(0, chartData.getNumOfSeries());
		Assert.assertEquals(null, chartData.getCategory(100)); //allow out of index
		
		chartData.setCategoriesFormula("A1:A3");
		Assert.assertEquals(3, chartData.getNumOfCategory());
		Assert.assertEquals("A", chartData.getCategory(0));
		Assert.assertEquals("B", chartData.getCategory(1));
		Assert.assertEquals("C", chartData.getCategory(2));
		
		SSeries nseries1 = chartData.addSeries();
		Assert.assertEquals(1, chartData.getNumOfSeries());
		Assert.assertEquals(null, nseries1.getName());
		
		nseries1.setXYFormula("KK()",null,null);//fail 
		Assert.assertEquals("#NAME?", nseries1.getName());
//		Assert.assertTrue(nseries1.isFormulaParsingError());
		
		nseries1.setXYFormula("D1",null,null);
		Assert.assertEquals("My Series", nseries1.getName());
		
		Assert.assertEquals(0, nseries1.getNumOfValue());
		Assert.assertEquals(0, nseries1.getNumOfYValue());
		Assert.assertFalse(nseries1.isFormulaParsingError());
		
		
		nseries1.setXYFormula("D1","KK()","KK()");
		Assert.assertEquals(0, nseries1.getNumOfValue());
		Assert.assertEquals(0, nseries1.getNumOfYValue());
//		Assert.assertTrue(nseries1.isFormulaParsingError());
		
		nseries1.setXYFormula("D1","B1:B3","C1:C3");
		Assert.assertFalse(nseries1.isFormulaParsingError());
		
		Assert.assertEquals(3, nseries1.getNumOfValue());
		Assert.assertEquals(3, nseries1.getNumOfYValue());
		
		Assert.assertEquals(1D, nseries1.getValue(0));
		Assert.assertEquals(2D, nseries1.getValue(1));
		Assert.assertEquals(3D, nseries1.getValue(2));
		
		Assert.assertEquals(4D, nseries1.getYValue(0));
		Assert.assertEquals(5D, nseries1.getYValue(1));
		Assert.assertEquals(6D, nseries1.getYValue(2));
		
		
		////
		SSeries nseries2 = chartData.addSeries();
		Assert.assertEquals(2, chartData.getNumOfSeries());
		Assert.assertEquals(null, nseries2.getName());
		
		Assert.assertEquals(0, nseries2.getNumOfValue());
		nseries2.setXYFormula(null,"C1:C3",null);
		
		Assert.assertEquals(3, nseries2.getNumOfValue());
		
		Assert.assertEquals(4D, nseries2.getValue(0));
		Assert.assertEquals(5D, nseries2.getValue(1));
		Assert.assertEquals(6D, nseries2.getValue(2));
	}
	
	@Test
	public void testDeleteRelease(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		SCell cell = sheet.getCell(10, 10);
		cell.setFormulaValue("SUM(999)");
		
		Assert.assertEquals(999D,cell.getNumberValue());
		
		book.deleteSheet(sheet);
		
		try{
			cell.getType();
			Assert.fail();
		}catch(IllegalStateException x){}
		
		
		try{
			cell.setValue("ABC");
			Assert.fail();
		}catch(IllegalStateException x){}
	}
	
	@Test
	public void testMergedRange(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		sheet.addMergedRegion(new CellRegion(1,1,2,2));
		sheet.addMergedRegion(new CellRegion(3,4,5,6));
		sheet.addMergedRegion(new CellRegion("J1:K2"));
		
		Assert.assertEquals(3, sheet.getMergedRegions().size());
		
		CellRegion region = sheet.getMergedRegions().get(0);
		Assert.assertEquals(1, region.row);
		Assert.assertEquals(1, region.column);
		Assert.assertEquals(2, region.lastRow);
		Assert.assertEquals(2, region.lastColumn);
		
		region = sheet.getMergedRegions().get(1);
		Assert.assertEquals(3, region.row);
		Assert.assertEquals(4, region.column);
		Assert.assertEquals(5, region.lastRow);
		Assert.assertEquals(6, region.lastColumn);
		
		region = sheet.getMergedRegions().get(2);
		Assert.assertEquals(0, region.row);
		Assert.assertEquals(9, region.column);
		Assert.assertEquals(1, region.lastRow);
		Assert.assertEquals(10, region.lastColumn);
		
		try{
			sheet.addMergedRegion(new CellRegion(0,0,1,1));
			Assert.fail();
		}catch(InvalidModelOpException x){}
		try{
			sheet.addMergedRegion(new CellRegion(1,1,2,2));
			Assert.fail();
		}catch(InvalidModelOpException x){}
		
		
		sheet = initialDataGrid(book.createSheet("Sheet 2"));
		
		sheet.addMergedRegion(new CellRegion(1,1,2,2));
		sheet.addMergedRegion(new CellRegion(1,7,2,8));
		sheet.addMergedRegion(new CellRegion(4,4,5,5));
		sheet.addMergedRegion(new CellRegion(7,1,8,2));
		sheet.addMergedRegion(new CellRegion(7,7,8,8));
		
		List<CellRegion> merges = sheet.getOverlapsMergedRegions(new CellRegion(1,2,8,3),false);
		
		Assert.assertEquals(2, merges.size());
		region = merges.get(0);
		Assert.assertEquals(1, region.row);
		Assert.assertEquals(1, region.column);
		Assert.assertEquals(2, region.lastRow);
		Assert.assertEquals(2, region.lastColumn);
		region = merges.get(1);
		Assert.assertEquals(7, region.row);
		Assert.assertEquals(1, region.column);
		Assert.assertEquals(8, region.lastRow);
		Assert.assertEquals(2, region.lastColumn);
		
		
		merges = sheet.getOverlapsMergedRegions(new CellRegion(1,2,8,4),false);
		
		Assert.assertEquals(3, merges.size());
		region = merges.get(0);
		Assert.assertEquals(1, region.row);
		Assert.assertEquals(1, region.column);
		Assert.assertEquals(2, region.lastRow);
		Assert.assertEquals(2, region.lastColumn);
		region = merges.get(1);
		Assert.assertEquals(4, region.row);
		Assert.assertEquals(4, region.column);
		Assert.assertEquals(5, region.lastRow);
		Assert.assertEquals(5, region.lastColumn);
		region = merges.get(2);
		Assert.assertEquals(7, region.row);
		Assert.assertEquals(1, region.column);
		Assert.assertEquals(8, region.lastRow);
		Assert.assertEquals(2, region.lastColumn);
		
		
		merges = sheet.getOverlapsMergedRegions(new CellRegion(2,2,5,5),false);
		
		Assert.assertEquals(2, merges.size());
		region = merges.get(0);
		Assert.assertEquals(1, region.row);
		Assert.assertEquals(1, region.column);
		Assert.assertEquals(2, region.lastRow);
		Assert.assertEquals(2, region.lastColumn);
		region = merges.get(1);
		Assert.assertEquals(4, region.row);
		Assert.assertEquals(4, region.column);
		Assert.assertEquals(5, region.lastRow);
		Assert.assertEquals(5, region.lastColumn);

		merges = sheet.getOverlapsMergedRegions(new CellRegion(2,2,6,6),false);
		
		Assert.assertEquals(2, merges.size());
		region = merges.get(0);
		Assert.assertEquals(1, region.row);
		Assert.assertEquals(1, region.column);
		Assert.assertEquals(2, region.lastRow);
		Assert.assertEquals(2, region.lastColumn);
		region = merges.get(1);
		Assert.assertEquals(4, region.row);
		Assert.assertEquals(4, region.column);
		Assert.assertEquals(5, region.lastRow);
		Assert.assertEquals(5, region.lastColumn);
		
	}
	@Test
	public void testSerializable(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet = initialDataGrid(book.createSheet("Sheet 1"));
		initialDataGrid(book.createSheet("Sheet 2"));
		Date now = new Date();
		
		sheet.getCell(1, 1).setStringValue("ABCD");
		sheet.getCell(2, 1).setupRichTextValue().addSegment("ABC", book.getDefaultFont());
		sheet.getCell(3, 1).setNumberValue(99D);
		sheet.getCell(4, 1).setDateValue(now);
		sheet.getCell(5, 1).setErrorValue(new ErrorValue(ErrorValue.INVALID_NAME));
		
		sheet.getCell(5, 1).setupHyperlink(HyperlinkType.URL,"httt://www.zkoss.org",null);
		
		sheet.getCell(5, 0).setupRichTextValue().addSegment("ABC",book.getDefaultFont());
		
		sheet.getCell(5, 1).setupComment().setText("AAA");
		sheet.getCell(5, 2).setupComment().setupRichText().addSegment("BBB",book.getDefaultFont());
		
		
		sheet.addMergedRegion(new CellRegion(0,1,2,3));
		
		SChart chart = sheet.addChart(ChartType.BAR, new ViewAnchor(0, 0, 800, 600));
		
		SGeneralChartData data = (SGeneralChartData)chart.getData();
		data.setCategoriesFormula("A1:A3");
		SSeries series = data.addSeries();
		series.setXYFormula("B1:B3", "C1:C3", null);
		
		sheet.addPicture(Format.PNG, new byte[]{}, new ViewAnchor(0, 0, 800, 600));
		
		SName name = book.createName("test");
		name.setRefersToFormula("'Sheet 1'!A1:B1");
		
		sheet.addDataValidation(new CellRegion("A1"));
		
		sheet.createAutoFilter(new CellRegion(0,0,20,20));
		
		ByteArrayOutputStream baos;
		ObjectOutputStream oos;
		try {
			baos = new ByteArrayOutputStream();
			oos = new ObjectOutputStream(baos);
			
			oos.writeObject(book);
			
			ObjectInputStream ois = new ObjectInputStream(new ByteArrayInputStream(baos.toByteArray()));
			
			book = (SBook)ois.readObject();

			sheet = book.getSheetByName("Sheet 1");
			
			Assert.assertNotNull(sheet);
			
			Assert.assertEquals("ABCD",sheet.getCell(1, 1).getStringValue());
			Assert.assertEquals("ABC",sheet.getCell(2, 1).getRichTextValue().getText());
			Assert.assertEquals(book.getDefaultFont(),sheet.getCell(2, 1).getRichTextValue().getSegments().get(0).getFont());
			Assert.assertEquals(99D,sheet.getCell(3, 1).getNumberValue());
			Assert.assertEquals(now,sheet.getCell(4, 1).getDateValue());
			Assert.assertEquals(ErrorValue.INVALID_NAME,sheet.getCell(5, 1).getErrorValue().getCode());
			
			Assert.assertEquals(HyperlinkType.URL,sheet.getCell(5, 1).getHyperlink().getType());
			
			Assert.assertEquals("ABC",sheet.getCell(5, 0).getRichTextValue().getText());
			
			Assert.assertEquals("AAA",sheet.getCell(5, 1).getComment().getText());
			Assert.assertEquals("BBB",sheet.getCell(5, 2).getComment().getRichText().getText());
			
			Assert.assertEquals(1, sheet.getMergedRegions().size());
			
			CellRegion region = sheet.getMergedRegions().get(0);
			Assert.assertEquals(0, region.row);
			Assert.assertEquals(1, region.column);
			Assert.assertEquals(2, region.lastRow);
			Assert.assertEquals(3, region.lastColumn);
			
			Assert.assertEquals(1, book.getNumOfName());
			name = book.getName(0);
			Assert.assertEquals("'Sheet 1'!A1:B1", name.getRefersToFormula());
			
			
			Assert.assertEquals(1, sheet.getCharts().size());
			chart = sheet.getCharts().get(0);
			
			data = (SGeneralChartData)chart.getData();
			
			Assert.assertEquals("A1:A3", data.getCategoriesFormula());
			Assert.assertEquals("B1:B3", data.getSeries(0).getNameFormula());
			Assert.assertEquals("C1:C3", data.getSeries(0).getValuesFormula());
			
			Assert.assertEquals(1, sheet.getPictures().size());
			SPicture picture = sheet.getPictures().get(0);
			
			Assert.assertEquals(1, sheet.getNumOfDataValidation());
			sheet.getDataValidation(0);
			
			Assert.assertEquals(new CellRegion(0,0,20,20), sheet.getAutoFilter().getRegion());
			
		} catch (Exception x) {
			throw new RuntimeException(x.getMessage(),x);
		}
		
		
	}
	
	@Test
	public void testName(){
		SBook book = SBooks.createBook("book1");
		initialDataGrid(book.createSheet("Sheet1"));
		
		SName name1 = book.createName("test1");
		try{
			book.createName("test1");
			Assert.fail();
		}catch(InvalidModelOpException e){}
		SName name2 = book.createName("test2");
		
		Assert.assertEquals(2, book.getNumOfName());
		Assert.assertEquals(name1, book.getName(0));
		Assert.assertEquals(name2, book.getName(1));
		
		Assert.assertNull(name1.getRefersToCellRegion());
		Assert.assertNull(name1.getRefersToSheetName());
		
		name1.setRefersToFormula("Sheet1!A1:B3");
		
		CellRegion region = name1.getRefersToCellRegion();
		Assert.assertEquals("Sheet1", name1.getRefersToSheetName());
		Assert.assertEquals("A1:B3", region.getReferenceString());
		Assert.assertEquals(0, region.row);
		Assert.assertEquals(0, region.column);
		Assert.assertEquals(2, region.lastRow);
		Assert.assertEquals(1, region.lastColumn);
		
		Assert.assertFalse(name2.isFormulaParsingError());
		
		name2.setRefersToFormula("Sheet2!A$2:B$4");
		
		region = name2.getRefersToCellRegion();
		Assert.assertEquals("Sheet2", name2.getRefersToSheetName());
		Assert.assertEquals("A2:B4", region.getReferenceString());
		Assert.assertEquals(1, region.row);
		Assert.assertEquals(0, region.column);
		Assert.assertEquals(3, region.lastRow);
		Assert.assertEquals(1, region.lastColumn);
		
		Assert.assertFalse(name2.isFormulaParsingError());
		name2.setRefersToFormula("IBM)(");
		Assert.assertTrue(name2.isFormulaParsingError());
		
	}
	
	@Test
	public void testStyleOptimal(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		SCellStyle defaultStyle = book.getDefaultCellStyle();
		SFont defaultFont = book.getDefaultFont();
		
		if(!(book instanceof BookImpl)){
			Assert.fail("not a book impl");
		}
		
		List<SCellStyle> styleTable = ((BookImpl)book).getCellStyleTable();
		List<SFont> fontTable = ((BookImpl)book).getFontTable();
		
		Assert.assertEquals(1, styleTable.size());
		Assert.assertEquals(defaultStyle, styleTable.get(0));
		
		Assert.assertEquals(1, fontTable.size());
		Assert.assertEquals(defaultFont, fontTable.get(0));
		
		book.createCellStyle(true).setAlignment(Alignment.LEFT);;//just cear , dosn't assign to cell
		book.createFont(true).setBoldweight(Boldweight.BOLD);
		
		Assert.assertEquals(2, styleTable.size());
		Assert.assertEquals(2, fontTable.size());
		
		book.optimizeCellStyle();
		Assert.assertEquals(1, styleTable.size());
		Assert.assertEquals(1, fontTable.size());
		
		SCellStyle style1,style2,style3,style4;
		SFont font1,font2;
		//style1 and style has same style but different font
		style1 = book.createCellStyle(true);
		style1.setAlignment(Alignment.LEFT);
		style2 = book.createCellStyle(true);
		style2.setAlignment(Alignment.LEFT);
		
		font1 = book.createFont(true);
		font1.setBoldweight(Boldweight.BOLD);
		style2.setFont(font1);
		
		
		//style 3 is same as default
		style3 = book.createCellStyle(true);
		
		//style 4 is same as default but has a font as default
		style4 = book.createCellStyle(true);
		font2 = book.createFont(true);
		style4.setFont(font2);
		
		sheet1.getCell(0, 0).setCellStyle(style1);
		sheet1.getCell(1, 1).setCellStyle(style2);
		sheet1.getCell(2, 2).setCellStyle(style3);
		sheet1.getCell(3, 3).setCellStyle(style4);
		
		Assert.assertEquals(5, styleTable.size());
		Assert.assertEquals(3, fontTable.size());
		
		book.optimizeCellStyle();
		Assert.assertEquals(3, styleTable.size());//
		Assert.assertEquals(2, fontTable.size());
		
		Assert.assertEquals(style1, sheet1.getCell(0, 0).getCellStyle());
		Assert.assertEquals(style2, sheet1.getCell(1, 1).getCellStyle());
		Assert.assertEquals(defaultStyle, sheet1.getCell(2, 2).getCellStyle());
		Assert.assertEquals(defaultStyle, sheet1.getCell(3, 3).getCellStyle());
	}
	
	@Test
	public void testDataValidation(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		sheet1.getCell(0, 0).setValue(1D);
		sheet1.getCell(0, 1).setValue(2D);
		sheet1.getCell(0, 2).setValue(3D);
		
		SDataValidation dv1 = sheet1.addDataValidation(new CellRegion(1,1));
		SDataValidation dv2 = sheet1.addDataValidation(new CellRegion(1,2));
		SDataValidation dv3 = sheet1.addDataValidation(new CellRegion(1,3));
		//LIST
		dv1.setValidationType(ValidationType.LIST);
		dv1.setFormula("A1:C1");
		
		Assert.assertEquals(3, dv1.getNumOfValue1());
		Assert.assertEquals(0, dv1.getNumOfValue2());
		Assert.assertEquals(1D, dv1.getValue1(0));
		Assert.assertEquals(2D, dv1.getValue1(1));
		Assert.assertEquals(3D, dv1.getValue1(2));
		
		
		dv2.setValidationType(ValidationType.INTEGER);
		dv2.setFormula("A1","C1");
		Assert.assertEquals(1, dv2.getNumOfValue1());
		Assert.assertEquals(1, dv2.getNumOfValue2());
		Assert.assertEquals(1D, dv2.getValue1(0));
		Assert.assertEquals(3D, dv2.getValue2(0));
		
		dv3.setValidationType(ValidationType.INTEGER);
		dv3.setFormula("AVERAGE(A1:C1)","SUM(A1:C1)");
		Assert.assertEquals(1, dv3.getNumOfValue1());
		Assert.assertEquals(1, dv3.getNumOfValue2());
		Assert.assertEquals(2D, dv3.getValue1(0));
		Assert.assertEquals(6D, dv3.getValue2(0));
		
		
		SRanges.range(sheet1,0,0).setEditText("2");
		SRanges.range(sheet1,0,1).setEditText("4");
		SRanges.range(sheet1,0,2).setEditText("6");
		
		Assert.assertEquals(3, dv1.getNumOfValue1());
		Assert.assertEquals(0, dv1.getNumOfValue2());
		Assert.assertEquals(2D, dv1.getValue1(0));
		Assert.assertEquals(4D, dv1.getValue1(1));
		Assert.assertEquals(6D, dv1.getValue1(2));
		
		Assert.assertEquals(1, dv2.getNumOfValue1());
		Assert.assertEquals(1, dv2.getNumOfValue2());
		Assert.assertEquals(2D, dv2.getValue1(0));
		Assert.assertEquals(6D, dv2.getValue2(0));
		
		Assert.assertEquals(1, dv3.getNumOfValue1());
		Assert.assertEquals(1, dv3.getNumOfValue2());
		Assert.assertEquals(4D, dv3.getValue1(0));
		Assert.assertEquals(12D, dv3.getValue2(0));
		
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		
		Set<Ref> refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 0)));
		Assert.assertEquals(3, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 1)));
		Assert.assertEquals(2, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 2)));
		Assert.assertEquals(3, refs.size());
		
		sheet1.deleteDataValidation(dv1);
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 0)));
		Assert.assertEquals(2, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 1)));
		Assert.assertEquals(1, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 2)));
		Assert.assertEquals(2, refs.size());
		
		sheet1.deleteDataValidation(dv2);
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 0)));
		Assert.assertEquals(1, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 1)));
		Assert.assertEquals(1, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 2)));
		Assert.assertEquals(1, refs.size());
		
		sheet1.deleteDataValidation(dv3);
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 0)));
		Assert.assertEquals(0, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 1)));
		Assert.assertEquals(0, refs.size());
		refs = table.getDependents(new RefImpl((AbstractCellAdv)sheet1.getCell(0, 2)));
		Assert.assertEquals(0, refs.size());
	}
	
//	@Test
//	public void testRowColumnInsertDeleteDependency(){
//		NBook book = NBooks.createBook("book1");
//		NSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
//		
//		sheet1.getCell("C3").setValue(3);
//		sheet1.getCell("A1").setFormulaValue("C3");
//		
//		Assert.assertEquals(3D, sheet1.getCell("A1").getValue());
//		
//		sheet1.insertRow(1,1);
//		
//		
//	}
	
	@Test
	public void testInsertDeleteCellVertical(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A1");
		sheet1.getCell("A2").setValue("A2");
		sheet1.getCell("A3").setValue("A3");
		sheet1.getCell("A4").setValue("A4");
		sheet1.getCell("B1").setValue("B1");
		sheet1.getCell("B2").setValue("B2");
		sheet1.getCell("B3").setValue("B3");
		sheet1.getCell("B4").setValue("B4");
		sheet1.getCell("C1").setValue("C1");
		sheet1.getCell("C2").setValue("C2");
		sheet1.getCell("C3").setValue("C3");
		sheet1.getCell("C4").setValue("C4");
		sheet1.getCell("D1").setValue("D1");
		sheet1.getCell("D2").setValue("D2");
		sheet1.getCell("D3").setValue("D3");
		sheet1.getCell("D4").setValue("D4");
		
		sheet1.insertCell(1, 1, 2, 2, false);
		
		Assert.assertEquals("A1",sheet1.getCell("A1").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("B1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B2",sheet1.getCell("B4").getValue());
		Assert.assertEquals("B3",sheet1.getCell("B5").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B6").getValue());
		
		Assert.assertEquals("C1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals(null,sheet1.getCell("C3").getValue());
		Assert.assertEquals("C2",sheet1.getCell("C4").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C5").getValue());
		Assert.assertEquals("C4",sheet1.getCell("C6").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("D2").getValue());
		Assert.assertEquals("D3",sheet1.getCell("D3").getValue());
		Assert.assertEquals("D4",sheet1.getCell("D4").getValue());
		
		
		sheet1.deleteCell(2, 2, 3, 3, false);
		
		Assert.assertEquals("A1",sheet1.getCell("A1").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("B1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B2",sheet1.getCell("B4").getValue());
		Assert.assertEquals("B3",sheet1.getCell("B5").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B6").getValue());
		
		Assert.assertEquals("C1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C3").getValue());
		Assert.assertEquals("C4",sheet1.getCell("C4").getValue());
		Assert.assertEquals(null,sheet1.getCell("C5").getValue());
		Assert.assertEquals(null,sheet1.getCell("C6").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("D2").getValue());
		Assert.assertEquals(null,sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		
		sheet1.insertCell(0, 0, 0, 0, false);
		Assert.assertEquals(null,sheet1.getCell("A1").getValue());
		Assert.assertEquals("A1",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A4").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A5").getValue());
		
		Assert.assertEquals("B1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B2",sheet1.getCell("B4").getValue());
		Assert.assertEquals("B3",sheet1.getCell("B5").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B6").getValue());
		
		Assert.assertEquals("C1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C3").getValue());
		Assert.assertEquals("C4",sheet1.getCell("C4").getValue());
		Assert.assertEquals(null,sheet1.getCell("C5").getValue());
		Assert.assertEquals(null,sheet1.getCell("C6").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("D2").getValue());
		Assert.assertEquals(null,sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		
		
		sheet1.deleteCell(0, 0, 0, 0, false);
		Assert.assertEquals("A1",sheet1.getCell("A1").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("B1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B2",sheet1.getCell("B4").getValue());
		Assert.assertEquals("B3",sheet1.getCell("B5").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B6").getValue());
		
		Assert.assertEquals("C1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C3").getValue());
		Assert.assertEquals("C4",sheet1.getCell("C4").getValue());
		Assert.assertEquals(null,sheet1.getCell("C5").getValue());
		Assert.assertEquals(null,sheet1.getCell("C6").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("D2").getValue());
		Assert.assertEquals(null,sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		
		
		sheet1.insertCell(100, 100, 101, 101, false);
		sheet1.deleteCell(100, 100, 101, 101, false);
		
	}
	
	@Test
	public void testInsertDeleteCellHorzontal(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		sheet1.getCell("A1").setValue("A1");
		sheet1.getCell("A2").setValue("A2");
		sheet1.getCell("A3").setValue("A3");
		sheet1.getCell("A4").setValue("A4");
		sheet1.getCell("B1").setValue("B1");
		sheet1.getCell("B2").setValue("B2");
		sheet1.getCell("B3").setValue("B3");
		sheet1.getCell("B4").setValue("B4");
		sheet1.getCell("C1").setValue("C1");
		sheet1.getCell("C2").setValue("C2");
		sheet1.getCell("C3").setValue("C3");
		sheet1.getCell("C4").setValue("C4");
		sheet1.getCell("D1").setValue("D1");
		sheet1.getCell("D2").setValue("D2");
		sheet1.getCell("D3").setValue("D3");
		sheet1.getCell("D4").setValue("D4");
		
		sheet1.insertCell(1, 1, 2, 2, true);
		
		Assert.assertEquals("A1",sheet1.getCell("A1").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("B1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B4").getValue());
		
		
		Assert.assertEquals("C1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals(null,sheet1.getCell("C3").getValue());
		Assert.assertEquals("C4",sheet1.getCell("C4").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("B2",sheet1.getCell("D2").getValue());
		Assert.assertEquals("B3",sheet1.getCell("D3").getValue());
		Assert.assertEquals("D4",sheet1.getCell("D4").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("E1").getValue());
		Assert.assertEquals("C2",sheet1.getCell("E2").getValue());
		Assert.assertEquals("C3",sheet1.getCell("E3").getValue());
		Assert.assertEquals(null,sheet1.getCell("E4").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("F1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("F2").getValue());
		Assert.assertEquals("D3",sheet1.getCell("F3").getValue());
		Assert.assertEquals(null,sheet1.getCell("F4").getValue());
		
		
		sheet1.deleteCell(2, 2, 3, 3, true);
		
		Assert.assertEquals("A1",sheet1.getCell("A1").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("B1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B4").getValue());
		
		
		Assert.assertEquals("C1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C3").getValue());
		Assert.assertEquals(null,sheet1.getCell("C4").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("B2",sheet1.getCell("D2").getValue());
		Assert.assertEquals("D3",sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("E1").getValue());
		Assert.assertEquals("C2",sheet1.getCell("E2").getValue());
		Assert.assertEquals(null,sheet1.getCell("E3").getValue());
		Assert.assertEquals(null,sheet1.getCell("E4").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("F1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("F2").getValue());
		Assert.assertEquals(null,sheet1.getCell("F3").getValue());
		Assert.assertEquals(null,sheet1.getCell("F4").getValue());
		
		
		sheet1.insertCell(0, 0, 0, 0, true);
		Assert.assertEquals(null,sheet1.getCell("A1").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("A1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B4").getValue());
		
		
		Assert.assertEquals("B1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C3").getValue());
		Assert.assertEquals(null,sheet1.getCell("C4").getValue());
		
		Assert.assertEquals("C1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("B2",sheet1.getCell("D2").getValue());
		Assert.assertEquals("D3",sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("E1").getValue());
		Assert.assertEquals("C2",sheet1.getCell("E2").getValue());
		Assert.assertEquals(null,sheet1.getCell("E3").getValue());
		Assert.assertEquals(null,sheet1.getCell("E4").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("F1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("F2").getValue());
		Assert.assertEquals(null,sheet1.getCell("F3").getValue());
		Assert.assertEquals(null,sheet1.getCell("F4").getValue());
		
		sheet1.deleteCell(0, 0, 0, 0, true);
		Assert.assertEquals("A1",sheet1.getCell("A1").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A2").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A3").getValue());
		Assert.assertEquals("A4",sheet1.getCell("A4").getValue());
		
		Assert.assertEquals("B1",sheet1.getCell("B1").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("B4",sheet1.getCell("B4").getValue());
		
		
		Assert.assertEquals("C1",sheet1.getCell("C1").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C3").getValue());
		Assert.assertEquals(null,sheet1.getCell("C4").getValue());
		
		Assert.assertEquals("D1",sheet1.getCell("D1").getValue());
		Assert.assertEquals("B2",sheet1.getCell("D2").getValue());
		Assert.assertEquals("D3",sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("E1").getValue());
		Assert.assertEquals("C2",sheet1.getCell("E2").getValue());
		Assert.assertEquals(null,sheet1.getCell("E3").getValue());
		Assert.assertEquals(null,sheet1.getCell("E4").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("F1").getValue());
		Assert.assertEquals("D2",sheet1.getCell("F2").getValue());
		Assert.assertEquals(null,sheet1.getCell("F3").getValue());
		Assert.assertEquals(null,sheet1.getCell("F4").getValue());
		
		sheet1.insertCell(100, 100, 101, 101, true);
		sheet1.deleteCell(100, 100, 101, 101, true);
	}
	
	@Test
	public void testInsertCellBlank(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		

		sheet1.getCell("A2").setValue("A2");
//		sheet1.getCell("B2").setValue("B2");
		sheet1.getCell("C2").setValue("C2");
		sheet1.getCell("A3").setValue("A3");
		sheet1.getCell("B3").setValue("B3");
		sheet1.getCell("C3").setValue("C3");
		
		sheet1.insertCell(0, 0, 0, 2, false);//row 1, A-C
		Assert.assertEquals(null,sheet1.getCell("A2").getValue());
		Assert.assertEquals(null,sheet1.getCell("B2").getValue());
		Assert.assertEquals(null,sheet1.getCell("C2").getValue());
		Assert.assertEquals("A2",sheet1.getCell("A3").getValue());
		Assert.assertEquals(null,sheet1.getCell("B3").getValue());
		Assert.assertEquals("C2",sheet1.getCell("C3").getValue());
		Assert.assertEquals("A3",sheet1.getCell("A4").getValue());
		Assert.assertEquals("B3",sheet1.getCell("B4").getValue());
		Assert.assertEquals("C3",sheet1.getCell("C4").getValue());
		
		
	}
	
	@Test
	public void testInsertExceed(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		int maxRow = book.getMaxRowSize();
		int maxColumn = book.getMaxColumnSize();
		
		SRow row0 = sheet1.getRow(maxRow-2);
		SRow row1 = sheet1.getRow(maxRow-1);
		SRow row2 = sheet1.getRow(maxRow);
		
		row0.setHeight(99);
		row1.setHeight(100);
		Assert.assertEquals(99, sheet1.getRow(maxRow-2).getHeight());
		Assert.assertEquals(100, sheet1.getRow(maxRow-1).getHeight());
		
		try{
			row2.setHeight(100);
			Assert.fail();
		}catch(IllegalStateException x){}
		
		sheet1.insertRow(0, 0);
		Assert.assertEquals(sheet1.getDefaultRowHeight(), sheet1.getRow(maxRow-2).getHeight());
		Assert.assertEquals(99, sheet1.getRow(maxRow-1).getHeight());
		Assert.assertEquals(sheet1.getDefaultRowHeight(), sheet1.getRow(maxRow).getHeight());
		
		
		/////////column
		SColumn column0 = sheet1.getColumn(maxColumn-2);
		SColumn column1 = sheet1.getColumn(maxColumn-1);
		SColumn column2 = sheet1.getColumn(maxColumn);
		
		column0.setWidth(33);
		column1.setWidth(55);
		Assert.assertEquals(33, sheet1.getColumn(maxColumn-2).getWidth());
		Assert.assertEquals(55, sheet1.getColumn(maxColumn-1).getWidth());
		
		try{
			column2.setWidth(100);
			Assert.fail();
		}catch(IllegalStateException x){}
		
		sheet1.insertColumn(0, 0);
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), sheet1.getColumn(maxColumn-2).getWidth());
		Assert.assertEquals(33, sheet1.getColumn(maxColumn-1).getWidth());
		Assert.assertEquals(sheet1.getDefaultColumnWidth(), sheet1.getColumn(maxColumn).getWidth());
		
		
		//cell vertical
		SCell cell0 = sheet1.getCell(maxRow-2,0);
		SCell cell1 = sheet1.getCell(maxRow-1,0);
		SCell cell2 = sheet1.getCell(maxRow,0);
		
		cell0.setValue("A");
		cell1.setValue("B");
		Assert.assertEquals("A", sheet1.getCell(maxRow-2,0).getValue());
		Assert.assertEquals("B", sheet1.getCell(maxRow-1,0).getValue());
		
		try{
			cell2.setValue("C");
			Assert.fail();
		}catch(IllegalStateException x){}
		
		sheet1.insertCell(0, 0, 0, 0,false);
		Assert.assertEquals(null, sheet1.getCell(maxRow-2,0).getValue());
		Assert.assertEquals("A", sheet1.getCell(maxRow-1,0).getValue());
		Assert.assertEquals(null, sheet1.getCell(maxRow,0).getValue());
		
		//cell horizontal
		cell0 = sheet1.getCell(0,maxColumn-2);
		cell1 = sheet1.getCell(0,maxColumn-1);
		cell2 = sheet1.getCell(0,maxColumn);
		
		cell0.setValue("X");
		cell1.setValue("Y");
		Assert.assertEquals("X", sheet1.getCell(0,maxColumn-2).getValue());
		Assert.assertEquals("Y", sheet1.getCell(0,maxColumn-1).getValue());
		
		try{
			cell2.setValue("C");
			Assert.fail();
		}catch(IllegalStateException x){}
		
		sheet1.insertCell(0, 0, 0, 0,true);
		Assert.assertEquals(null, sheet1.getCell(0,maxColumn-2).getValue());
		Assert.assertEquals("X", sheet1.getCell(0,maxColumn-1).getValue());
		Assert.assertEquals(null, sheet1.getCell(0,maxColumn).getValue());		
	}
	
	@Test
	public void testMoveCell(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		
		sheet1.getCell("D3").setValue("D3");
		sheet1.getCell("D4").setValue("D4");
		sheet1.getCell("D5").setValue("D5");
		sheet1.getCell("E3").setValue("E3");
//		sheet1.getCell("E4").setValue("E4");//E4 is empty
		sheet1.getCell("E5").setValue("E5");
		sheet1.getCell("F3").setValue("F3");
		sheet1.getCell("F4").setValue("F4");
		sheet1.getCell("F5").setValue("F5");
		
		System.out.println(">>>>>>>>>>>>> Down");
		//fill  data that will be replaced when move
		sheet1.getCell("D6").setValue("A");
		sheet1.getCell("E6").setValue("B");
		sheet1.getCell("F6").setValue("C");
		sheet1.moveCell(2, 3, 4, 5, 1, 0);
		
		Assert.assertEquals(null,sheet1.getCell("D3").getValue());
		Assert.assertEquals("D3",sheet1.getCell("D4").getValue());
		Assert.assertEquals("D4",sheet1.getCell("D5").getValue());
		Assert.assertEquals("D5",sheet1.getCell("D6").getValue());
		Assert.assertEquals(null,sheet1.getCell("D7").getValue());
		Assert.assertEquals(null,sheet1.getCell("E3").getValue());
		Assert.assertEquals("E3",sheet1.getCell("E4").getValue());
		Assert.assertEquals(null,sheet1.getCell("E5").getValue());
		Assert.assertEquals("E5",sheet1.getCell("E6").getValue());
		Assert.assertEquals(null,sheet1.getCell("E7").getValue());
		Assert.assertEquals(null,sheet1.getCell("F3").getValue());
		Assert.assertEquals("F3",sheet1.getCell("F4").getValue());
		Assert.assertEquals("F4",sheet1.getCell("F5").getValue());
		Assert.assertEquals("F5",sheet1.getCell("F6").getValue());
		Assert.assertEquals(null,sheet1.getCell("F7").getValue());
		
		System.out.println(">>>>>>>>>>>>> Up");
		//fill  data that will be replaced when move
		sheet1.getCell("D3").setValue("A");
		sheet1.getCell("E3").setValue("B");
		sheet1.getCell("F3").setValue("C");
		sheet1.moveCell(3, 3, 5, 5, -1, 0);
		Assert.assertEquals(null,sheet1.getCell("D2").getValue());
		Assert.assertEquals("D3",sheet1.getCell("D3").getValue());
		Assert.assertEquals("D4",sheet1.getCell("D4").getValue());
		Assert.assertEquals("D5",sheet1.getCell("D5").getValue());
		Assert.assertEquals(null,sheet1.getCell("E2").getValue());
		Assert.assertEquals("E3",sheet1.getCell("E3").getValue());
		Assert.assertEquals(null,sheet1.getCell("E4").getValue());
		Assert.assertEquals("E5",sheet1.getCell("E5").getValue());
		Assert.assertEquals(null,sheet1.getCell("F2").getValue());
		Assert.assertEquals("F3",sheet1.getCell("F3").getValue());
		Assert.assertEquals("F4",sheet1.getCell("F4").getValue());
		Assert.assertEquals("F5",sheet1.getCell("F5").getValue());
		
		System.out.println(">>>>>>>>>>>>> Right");
		//fill  data that will be replaced when move
		sheet1.getCell("G3").setValue("A");
		sheet1.getCell("G4").setValue("B");
		sheet1.getCell("G5").setValue("C");		
		sheet1.moveCell(2, 3, 4, 5, 0, 1);
		Assert.assertEquals(null,sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		Assert.assertEquals(null,sheet1.getCell("D5").getValue());
		Assert.assertEquals("D3",sheet1.getCell("E3").getValue());
		Assert.assertEquals("D4",sheet1.getCell("E4").getValue());
		Assert.assertEquals("D5",sheet1.getCell("E5").getValue());
		Assert.assertEquals("E3",sheet1.getCell("F3").getValue());
		Assert.assertEquals(null,sheet1.getCell("F4").getValue());
		Assert.assertEquals("E5",sheet1.getCell("F5").getValue());
		Assert.assertEquals("F3",sheet1.getCell("G3").getValue());
		Assert.assertEquals("F4",sheet1.getCell("G4").getValue());
		Assert.assertEquals("F5",sheet1.getCell("G5").getValue());
		Assert.assertEquals(null,sheet1.getCell("H3").getValue());
		Assert.assertEquals(null,sheet1.getCell("H4").getValue());
		Assert.assertEquals(null,sheet1.getCell("H5").getValue());
		
		System.out.println(">>>>>>>>>>>>> Left");
		//fill  data that will be replaced when move
		sheet1.getCell("D3").setValue("A");
		sheet1.getCell("D4").setValue("B");
		sheet1.getCell("D5").setValue("C");			
		sheet1.moveCell(2, 4, 4, 6, 0, -1);
		Assert.assertEquals(null,sheet1.getCell("C3").getValue());
		Assert.assertEquals(null,sheet1.getCell("C4").getValue());
		Assert.assertEquals(null,sheet1.getCell("C5").getValue());
		Assert.assertEquals("D3",sheet1.getCell("D3").getValue());
		Assert.assertEquals("D4",sheet1.getCell("D4").getValue());
		Assert.assertEquals("D5",sheet1.getCell("D5").getValue());
		Assert.assertEquals("E3",sheet1.getCell("E3").getValue());
		Assert.assertEquals(null,sheet1.getCell("E4").getValue());
		Assert.assertEquals("E5",sheet1.getCell("E5").getValue());
		Assert.assertEquals("F3",sheet1.getCell("F3").getValue());
		Assert.assertEquals("F4",sheet1.getCell("F4").getValue());
		Assert.assertEquals("F5",sheet1.getCell("F5").getValue());
		
		System.out.println(">>>>>>>>>>>>> Down-Right");
		sheet1.moveCell(2, 3, 4, 5, 1, 1);
		Assert.assertEquals(null,sheet1.getCell("D3").getValue());
		Assert.assertEquals(null,sheet1.getCell("D4").getValue());
		Assert.assertEquals(null,sheet1.getCell("D5").getValue());
		Assert.assertEquals(null,sheet1.getCell("D6").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("E3").getValue());
		Assert.assertEquals("D3",sheet1.getCell("E4").getValue());
		Assert.assertEquals("D4",sheet1.getCell("E5").getValue());
		Assert.assertEquals("D5",sheet1.getCell("E6").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("F3").getValue());
		Assert.assertEquals("E3",sheet1.getCell("F4").getValue());
		Assert.assertEquals(null,sheet1.getCell("F5").getValue());
		Assert.assertEquals("E5",sheet1.getCell("F6").getValue());
		
		Assert.assertEquals(null,sheet1.getCell("G3").getValue());
		Assert.assertEquals("F3",sheet1.getCell("G4").getValue());
		Assert.assertEquals("F4",sheet1.getCell("G5").getValue());
		Assert.assertEquals("F5",sheet1.getCell("G6").getValue());
		
		
		System.out.println(">>>>>>>>>>>>> Top-Left");
		sheet1.moveCell(3, 4, 5, 6, -1, -1);
		Assert.assertEquals("D3",sheet1.getCell("D3").getValue());
		Assert.assertEquals("D4",sheet1.getCell("D4").getValue());
		Assert.assertEquals("D5",sheet1.getCell("D5").getValue());
		Assert.assertEquals("E3",sheet1.getCell("E3").getValue());
		Assert.assertEquals(null,sheet1.getCell("E4").getValue());
		Assert.assertEquals("E5",sheet1.getCell("E5").getValue());
		Assert.assertEquals("F3",sheet1.getCell("F3").getValue());
		Assert.assertEquals("F4",sheet1.getCell("F4").getValue());
		Assert.assertEquals("F5",sheet1.getCell("F5").getValue());
	}
	
	@Test
	public void testMoveCellWithMerge(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		sheet1.getCell("A1").setValue("A");
		sheet1.addMergedRegion(new CellRegion("B2:C3"));
		sheet1.addMergedRegion(new CellRegion("D2:E3"));
		
		try{
			sheet1.moveCell(new CellRegion("A1:C2"), 1,1);//source overlap
			Assert.fail();
		}catch(InvalidModelOpException x){}
		
		
		sheet1.moveCell(new CellRegion("A1:C3"), 0,3);//target overlap
		//it unmerge the target and then move source
		Assert.assertEquals(1, sheet1.getNumOfMergedRegion());
		Assert.assertEquals("E2:F3", sheet1.getMergedRegions().get(0).getReferenceString());
		

		
		sheet1.moveCell(new CellRegion("D1:F3"), 4, -2);
		
		Assert.assertEquals("A", sheet1.getCell("B5").getValue());
		Assert.assertEquals(1, sheet1.getNumOfMergedRegion());
		Assert.assertEquals("C6:D7", sheet1.getMergedRegion(5, 2).getReferenceString());
		Assert.assertEquals("C6:D7", sheet1.getMergedRegion("C6").getReferenceString());
		Assert.assertEquals("C6:D7", sheet1.getMergedRegion("C7").getReferenceString());
		Assert.assertEquals("C6:D7", sheet1.getMergedRegion("D6").getReferenceString());
		Assert.assertEquals("C6:D7", sheet1.getMergedRegion("D7").getReferenceString());
		
		
		
	}
	
	@Test
	public void testAutoFilter(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		
		
		Assert.assertNull(sheet1.getAutoFilter());
		
		SAutoFilter filter = sheet1.createAutoFilter(new CellRegion(0,0,1,1));
		Assert.assertNotNull(filter);
		Assert.assertEquals(filter, sheet1.getAutoFilter());
		
		
		Assert.assertNull(filter.getFilterColumn(0, false));
		NFilterColumn col0 = filter.getFilterColumn(0, true);
		Assert.assertEquals(col0, filter.getFilterColumn(0, true));
		
		col0.setProperties(FilterOp.VALUES, new String[]{"ABC","DEF"}, null, false);

		
		NFilterColumn col1 = filter.getFilterColumn(1, true);
		try{
			filter.getFilterColumn(2, true);	
			Assert.fail();
		}catch(IllegalStateException x){}
		
		Assert.assertEquals(FilterOp.VALUES, col0.getOperator());
		Assert.assertEquals(2,col0.getCriteria1().size());
		Assert.assertEquals("ABC",col0.getCriteria1().iterator().next());
		Assert.assertEquals(0,col0.getCriteria2().size());
		Assert.assertEquals(2,col0.getFilters().size());
		Assert.assertEquals("ABC",col0.getFilters().get(0));
		Assert.assertEquals("DEF",col0.getFilters().get(1));
		
		sheet1.clearAutoFilter();
		Assert.assertNull(sheet1.getAutoFilter());
		
	}
	
	@Test
	public void testDelRowAndShrinkMerged() {
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Sheet1");
		AbstractSheetAdv sheetAdv = (AbstractSheetAdv)sheet;

		CellRegion mergedCell = new CellRegion(3, 0, 5, 10);
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 0, 2, new CellRegion(0, 0, 2, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 1, 2, new CellRegion(1, 0, 3, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 2, 3, new CellRegion(2, 0, 3, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 2, 4, new CellRegion(2, 0, 2, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 3, 5, null);
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 4, 4, new CellRegion(3, 0, 4, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 4, 5, new CellRegion(3, 0, 3, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 5, 5, new CellRegion(3, 0, 4, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 4, 6, new CellRegion(3, 0, 3, 10));
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 2, 5, null);
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 2, 6, null);
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 3, 5, null);
		testDelRowAndShrinkMerged(sheetAdv, mergedCell, 3, 6, null);
	}
	
	private void testDelRowAndShrinkMerged(AbstractSheetAdv sheet, CellRegion mergedCell, int row, int lastRow, CellRegion expected) {
		sheet.addMergedRegion(mergedCell);
		sheet.deleteRow(row, lastRow);
		if(expected != null) {
			Assert.assertEquals(sheet.getMergedRegion(0), expected);
			sheet.removeMergedRegion(expected, true);
		} else {
			Assert.assertTrue(sheet.getMergedRegions().isEmpty());
		}
	}
	
	@Test
	public void testDelColAndShrinkMerged() {
		SBook book = SBooks.createBook("book1");
		SSheet sheet = book.createSheet("Sheet1");
		AbstractSheetAdv sheetAdv = (AbstractSheetAdv)sheet;

		CellRegion mergedCell = new CellRegion(0, 3, 10, 5);
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 0, 2, new CellRegion(0, 0, 10, 2));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 1, 2, new CellRegion(0, 1, 10, 3));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 2, 3, new CellRegion(0, 2, 10, 3));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 2, 4, new CellRegion(0, 2, 10, 2));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 3, 5, null);
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 4, 4, new CellRegion(0, 3, 10, 4));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 4, 5, new CellRegion(0, 3, 10, 3));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 5, 5, new CellRegion(0, 3, 10, 4));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 4, 6, new CellRegion(0, 3, 10, 3));
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 2, 5, null);
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 2, 6, null);
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 3, 5, null);
		testDelColAndShrinkMerged(sheetAdv, mergedCell, 3, 6, null);
	}
	
	private void testDelColAndShrinkMerged(AbstractSheetAdv sheet, CellRegion mergedCell, int col, int lastCol, CellRegion expected) {
		sheet.addMergedRegion(mergedCell);
		sheet.deleteColumn(col, lastCol);
		if(expected != null) {
			Assert.assertEquals(sheet.getMergedRegion(0), expected);
			sheet.removeMergedRegion(expected, true);
		} else {
			Assert.assertTrue(sheet.getMergedRegions().isEmpty());
		}
	}
	
	@Test
	public void testMultipleAreaEval() {
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		book.createName("SingleValueName").setRefersToFormula("Sheet1!$A$1");
		book.createName("SingleName").setRefersToFormula("Sheet1!$A$1:$A$3");
		book.createName("MultipleName").setRefersToFormula("Sheet1!$A$1,Sheet1!$A$3");
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("A2").setValue(2);
		sheet1.getCell("A3").setValue(3);
		sheet2.getCell("A5").setValue(5);
		sheet2.getCell("A6").setValue(6);
		
		sheet1.getCell("D2").setValue("=SUM(A1,A3,Sheet2!A5:A6)");
		Assert.assertEquals(15D, sheet1.getCell("D2").getValue());
		
		sheet1.getCell("D3").setValue("=(A1,A3,Sheet2!A5:A6)");
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());
		
		sheet1.getCell("D4").setValue("=SUM(SingleValueName)");
		Assert.assertEquals(1D, sheet1.getCell("D4").getValue());
		
		sheet1.getCell("D5").setValue("=SingleValueName");
		Assert.assertEquals(1D, sheet1.getCell("D5").getValue());
		
		sheet1.getCell("D6").setValue("=SUM(SingleName)");
		Assert.assertEquals(6D, sheet1.getCell("D6").getValue());
		
		
		//the poi parser can handle =SUM(MultipleName) and get java.lang.IllegalStateException: evaluation stack not empty
//		sheet1.getCell("D7").setValue("=SUM(MultipleName)");
//		Assert.assertEquals(4D, sheet1.getCell("D7").getValue());
		
		//the poi parser can handle =SUM(MultipleName) and get java.lang.IllegalStateException: evaluation stack not empty
//		sheet1.getCell("D8").setValue("=MultipleName");
//		Assert.assertEquals("#VALUE!", sheet1.getCell("D8").getErrorValue().getErrorString());
	}
	@Test
	public void testMultipleAreaEvalOfLoop() {
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		book.createName("SingleValueName").setRefersToFormula("Sheet1!$A$1");
		book.createName("SingleName").setRefersToFormula("Sheet1!$A$1:$A$3");
		book.createName("MultipleName").setRefersToFormula("Sheet1!$A$1,Sheet1!$A$3");
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("A2").setValue(2);
		sheet1.getCell("A3").setValue(3);
		sheet2.getCell("A5").setValue(5);
		sheet2.getCell("A6").setValue(6);
		sheet1.getCell("B1").setValue("A");
		sheet1.getCell("B3").setValue("B");		
		
		//loop
		sheet1.getCell("C1").setValue("=C2");
		sheet1.getCell("C2").setValue("=C3");
		sheet1.getCell("C3").setValue("=C1");
		
		Assert.assertEquals("#N/A", sheet1.getCell("C1").getErrorValue().getErrorString());
		Assert.assertEquals("#N/A", sheet1.getCell("C2").getErrorValue().getErrorString());
		Assert.assertEquals("#N/A", sheet1.getCell("C3").getErrorValue().getErrorString());
		
		sheet1.getCell("D2").setValue("=SUM(A1,D2,Sheet2!A5:A6)");
		Assert.assertEquals("#N/A", sheet1.getCell("D2").getErrorValue().getErrorString());
		
		sheet1.getCell("D3").setValue("=(A1,D3,Sheet2!A5:A6)");
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());
		
		
		SChart p1 = sheet1.addChart(SChart.ChartType.PIE, new ViewAnchor(6,6, 600,400));
		
		SGeneralChartData data = (SGeneralChartData)p1.getData();
		data.setCategoriesFormula("(Sheet1!$B$1,Sheet1!$B$3)");
		
		Assert.assertEquals(2, data.getNumOfCategory());
		Assert.assertEquals("A", data.getCategory(0));
		Assert.assertEquals("B", data.getCategory(1));
		
		
		SSeries series = data.addSeries();
		series.setFormula(null, "(C1,C3)");
		Assert.assertEquals(0, series.getNumOfValue());
//		Assert.assertEquals(0D, series.getValue(0));
//		Assert.assertEquals(0D, series.getValue(1));
		
	}
	@Test
	public void testMultipleAreaShift() {
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		book.createName("SingleValueName").setRefersToFormula("Sheet1!$A$1");
		book.createName("SingleName").setRefersToFormula("Sheet1!$A$1:$A$3");
		book.createName("MultipleName").setRefersToFormula("Sheet1!$A$1,Sheet1!$A$3");
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("A2").setValue(2);
		sheet1.getCell("A3").setValue(3);
		
		sheet2.getCell("A5").setValue(5);
		sheet2.getCell("A6").setValue(6);
		
		sheet1.getCell("D2").setValue("=SUM(A1,A3,Sheet2!A5:A6)");
		Assert.assertEquals(15D, sheet1.getCell("D2").getValue());
		
		sheet1.getCell("D3").setValue("=(A1,A3,Sheet2!A5:A6)");
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());
		
		//shift
		SRanges.range(sheet1,"D2:D3").copy(SRanges.range(sheet1,"D22"));
		Assert.assertEquals("SUM(A21,A23,Sheet2!A25:A26)", sheet1.getCell("D22").getFormulaValue());
		sheet1.getCell("A21").setValue(3);
		sheet1.getCell("A23").setValue(6);
		sheet2.getCell("A25").setValue(1);
		sheet2.getCell("A26").setValue(2);
		Assert.assertEquals(12D, sheet1.getCell("D22").getValue());
		
		Assert.assertEquals("(A21,A23,Sheet2!A25:A26)", sheet1.getCell("D23").getFormulaValue());
		Assert.assertEquals("#VALUE!", sheet1.getCell("D23").getErrorValue().getErrorString());
		
		//move
		SRanges.range(sheet1,"A1:A3").move(2, 0);		
		
		Assert.assertEquals("SUM(A3,A5,Sheet2!A5:A6)", sheet1.getCell("D2").getFormulaValue());
		sheet1.getCell("A3").setValue(3);
		Assert.assertEquals(17D, sheet1.getCell("D2").getValue());
		
		Assert.assertEquals("(A3,A5,Sheet2!A5:A6)", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());
		
		//insert
		SRanges.range(sheet1,"A4").getRows().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);		
		
		Assert.assertEquals("SUM(A3,A6,Sheet2!A5:A6)", sheet1.getCell("D2").getFormulaValue());
		sheet1.getCell("A4").setValue(3);
		Assert.assertEquals(17D, sheet1.getCell("D2").getValue());
		sheet1.getCell("A6").setValue(5);
		Assert.assertEquals(19D, sheet1.getCell("D2").getValue());
		
		Assert.assertEquals("(A3,A6,Sheet2!A5:A6)", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());
		
		
		//delete
		SRanges.range(sheet1,"A4:A5").getRows().delete(DeleteShift.DEFAULT);		
		
		Assert.assertEquals("SUM(A3,A4,Sheet2!A5:A6)", sheet1.getCell("D2").getFormulaValue());
		sheet1.getCell("A4").setValue(3);
		Assert.assertEquals(17D, sheet1.getCell("D2").getValue());
		
		Assert.assertEquals("(A3,A4,Sheet2!A5:A6)", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());
		
		//rename
		SRanges.range(sheet2).setSheetName("XYZ");
		
		Assert.assertEquals("SUM(A3,A4,XYZ!A5:A6)", sheet1.getCell("D2").getFormulaValue());
		sheet2.getCell("A5").setValue(1);
		Assert.assertEquals(13D, sheet1.getCell("D2").getValue());
		
		Assert.assertEquals("(A3,A4,XYZ!A5:A6)", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());
		
		//delete
		SRanges.range(sheet2).deleteSheet();
		
		Assert.assertEquals("SUM(A3,A4,'#REF'!A5:A6)", sheet1.getCell("D2").getFormulaValue());
		Assert.assertEquals("#REF!", sheet1.getCell("D2").getErrorValue().getErrorString());
		
		Assert.assertEquals("(A3,A4,'#REF'!A5:A6)", sheet1.getCell("D3").getFormulaValue());
		Assert.assertEquals("#VALUE!", sheet1.getCell("D3").getErrorValue().getErrorString());		
	}
	
	@Test
	public void testMultipleAreaEvalOfChart() {
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		book.createName("SingleValueName").setRefersToFormula("Sheet1!$A$1");
		book.createName("SingleName").setRefersToFormula("Sheet1!$A$1:$A$3");
		book.createName("MultipleName").setRefersToFormula("Sheet1!$A$1,Sheet1!$A$3");
		sheet1.getCell("A1").setValue(1);
		sheet1.getCell("A2").setValue(2);
		sheet1.getCell("A3").setValue(3);
		sheet2.getCell("A5").setValue(5);
		sheet2.getCell("A6").setValue(6);
		
		sheet1.getCell("B1").setValue("A");
		sheet1.getCell("B3").setValue("B");
	
		SChart p1 = sheet1.addChart(SChart.ChartType.PIE, new ViewAnchor(6,6, 600,400));
		
		SGeneralChartData data = (SGeneralChartData)p1.getData();
		data.setCategoriesFormula("(Sheet1!$B$1,Sheet1!$B$3)");
		
		Assert.assertEquals(2, data.getNumOfCategory());
		Assert.assertEquals("A", data.getCategory(0));
		Assert.assertEquals("B", data.getCategory(1));
		
		
		SSeries series = data.addSeries();
		series.setFormula(null, "(A1,A3)");
		Assert.assertEquals(2, series.getNumOfValue());
		Assert.assertEquals(1D, series.getValue(0));
		Assert.assertEquals(3D, series.getValue(1));
		
		
		
	}
	
	@Test
	public void testCellRegionDiff() {
		CellRegion r1 = new CellRegion("C3:F6");
		
		CellRegion rightBottom = new CellRegion("E4:H7");
		CellRegion[] rr1 = {new CellRegion("C3:F3"), new CellRegion("C4:D6")};
		List<CellRegion> cc1 = r1.diff(rightBottom);
		testDiff_0(rr1, cc1.toArray(new CellRegion[cc1.size()]));
		
		CellRegion rightTop = new CellRegion("E2:H3");
		CellRegion[] rr2 = {new CellRegion("C4:F6"), new CellRegion("C3:D3")};
		List<CellRegion> cc2 = r1.diff(rightTop);
		testDiff_0(rr2, cc2.toArray(new CellRegion[cc2.size()]));
		
		CellRegion leftTop = new CellRegion("B2:E4");
		CellRegion[] rr3 = {new CellRegion("C5:F6"), new CellRegion("F3:F4")};
		List<CellRegion> cc3 = r1.diff(leftTop);
		testDiff_0(rr3, cc3.toArray(new CellRegion[cc3.size()]));
		
		CellRegion leftBottom = new CellRegion("B6:E7");
		CellRegion[] rr4 = {new CellRegion("C3:F5"), new CellRegion("F6")};
		List<CellRegion> cc4 = r1.diff(leftBottom);
		testDiff_0(rr4, cc4.toArray(new CellRegion[cc4.size()]));	
		
		CellRegion crossMiddleVertical = new CellRegion("E1:E8");
		CellRegion[] rr5 = {new CellRegion("C3:D6"), new CellRegion("F3:F6")};
		List<CellRegion> cc5 = r1.diff(crossMiddleVertical);
		testDiff_0(rr5, cc5.toArray(new CellRegion[cc5.size()]));
		
		CellRegion crossMiddleHorizontal = new CellRegion("A5:G5");
		CellRegion[] rr6 = {new CellRegion("C3:F4"), new CellRegion("C6:F6")};
		List<CellRegion> cc6 = r1.diff(crossMiddleHorizontal);
		testDiff_0(rr6, cc6.toArray(new CellRegion[cc6.size()]));
		
		CellRegion top = new CellRegion("C2:F3");
		CellRegion[] rr7 = {new CellRegion("C4:F6")};
		List<CellRegion> cc7 = r1.diff(top);
		testDiff_0(rr7, cc7.toArray(new CellRegion[cc7.size()]));
		
		CellRegion right = new CellRegion("F3:G6");
		CellRegion[] rr8 = {new CellRegion("C3:E6")};
		List<CellRegion> cc8 = r1.diff(right);
		testDiff_0(rr8, cc8.toArray(new CellRegion[cc8.size()]));
		
		CellRegion center = new CellRegion("D4:E5");
		CellRegion[] rr9 = {new CellRegion("C3:F3"), new CellRegion("C6:F6"), new CellRegion("C4:C5"), new CellRegion("F4:F5")};
		List<CellRegion> cc9 = r1.diff(center);
		testDiff_0(rr9, cc9.toArray(new CellRegion[cc9.size()]));
	}
	
	private void testDiff_0(CellRegion[] expected, CellRegion[] actual) {
		
		Assert.assertEquals(expected.length, actual.length);
		
		for(int i = 0; i < actual.length; i++) {
			Assert.assertEquals(expected[i], actual[i]);
		}
	}
}
