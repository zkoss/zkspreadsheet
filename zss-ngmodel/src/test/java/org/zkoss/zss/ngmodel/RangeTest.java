package org.zkoss.zss.ngmodel;

import java.text.SimpleDateFormat;
import java.util.Date;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRanges;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.impl.BookImpl;

public class RangeTest {

	@Test
	public void testGetRange(){
		
		NBook book = new BookImpl();
		NSheet sheet1 = book.createSheet("Sheet1");
		
		
		NRange r1 = NRanges.range(sheet1);
		Assert.assertEquals(0, r1.getRow());
		Assert.assertEquals(0, r1.getColumn());
		Assert.assertEquals(book.getMaxRowSize(), r1.getLastRow());
		Assert.assertEquals(book.getMaxColumnSize(), r1.getLastColumn());
		
		
		r1 = NRanges.range(sheet1,3,4);
		Assert.assertEquals(3, r1.getRow());
		Assert.assertEquals(4, r1.getColumn());
		Assert.assertEquals(3, r1.getLastRow());
		Assert.assertEquals(4, r1.getLastColumn());
		
		r1 = NRanges.range(sheet1,3,4,5,6);
		Assert.assertEquals(3, r1.getRow());
		Assert.assertEquals(4, r1.getColumn());
		Assert.assertEquals(5, r1.getLastRow());
		Assert.assertEquals(6, r1.getLastColumn());
	}
	
	
	@Test
	public void testGeneralCellValue1(){
		NBook book = new BookImpl();
		NSheet sheet = book.createSheet("Sheet 1");
		Date now = new Date();
		ErrorValue err = new ErrorValue((byte)0);
		NCell cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		NRanges.range(sheet,1,1).setEditText("abc");
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("abc",cell.getValue());
		
		NRanges.range(sheet,1,1).setEditText("123");
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(123,cell.getNumberValue().intValue());
		
		NRanges.range(sheet,1,1).setEditText("2013/01/01");
		Assert.assertEquals(CellType.DATE, cell.getType());
		Assert.assertEquals("2013/01/01",new SimpleDateFormat("yyyy/MM/dd").format((Date)cell.getValue()));
		
		
		NRanges.range(sheet,1,1).setEditText("=SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999, cell.getValue());
		
		NRanges.range(sheet,1,1).setEditText("=SUM)((999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.ERROR, cell.getFormulaResultType());
		Assert.assertEquals("SUM)((999)", cell.getFormulaValue());
		Assert.assertTrue(cell.getValue() instanceof ErrorValue);
		
		NRanges.range(sheet,1,1).setEditText("");
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());		
		
	}
	
	@Test
	public void testGeneralCellValue2(){
		NBook book = new BookImpl();
		NSheet sheet = book.createSheet("Sheet 1");
		Date now = new Date();
		ErrorValue err = new ErrorValue((byte)0);
		NCell cell = sheet.getCell(1, 1);
		
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());
		
		NRanges.range(sheet,1,1).setValue("abc");
		Assert.assertEquals(CellType.STRING, cell.getType());
		Assert.assertEquals("abc",cell.getValue());
		
		NRanges.range(sheet,1,1).setValue(123);
		Assert.assertEquals(CellType.NUMBER, cell.getType());
		Assert.assertEquals(123,cell.getValue());
		
		
		NRanges.range(sheet,1,1).setValue(now);
		Assert.assertEquals(CellType.DATE, cell.getType());
		Assert.assertEquals(now,cell.getValue());
		
		
		NRanges.range(sheet,1,1).setValue("=SUM(999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.NUMBER, cell.getFormulaResultType());
		Assert.assertEquals("SUM(999)", cell.getFormulaValue());
		Assert.assertEquals(999, cell.getValue());
		
		NRanges.range(sheet,1,1).setValue("=SUM)((999)");
		Assert.assertEquals(CellType.FORMULA, cell.getType());
		Assert.assertEquals(CellType.ERROR, cell.getFormulaResultType());
		Assert.assertEquals("SUM)((999)", cell.getFormulaValue());
		Assert.assertTrue(cell.getValue() instanceof ErrorValue);
		
		NRanges.range(sheet,1,1).setValue("");
		Assert.assertEquals(CellType.BLANK, cell.getType());
		Assert.assertNull(cell.getValue());		
		
	}
}
