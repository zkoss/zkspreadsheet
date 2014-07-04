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
import org.junit.Test;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SCellStyle.Alignment;
import org.zkoss.zss.model.SCellStyle.BorderType;
import org.zkoss.zss.model.SCellStyle.FillPattern;
import org.zkoss.zss.model.SChart.ChartType;
import org.zkoss.zss.model.SDataValidation.OperatorType;
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
import org.zkoss.zss.model.impl.AbstractDataValidationAdv;
import org.zkoss.zss.model.impl.AbstractSheetAdv;
import org.zkoss.zss.model.impl.BookImpl;
import org.zkoss.zss.model.impl.RefImpl;
import org.zkoss.zss.model.impl.SheetImpl;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.util.CellStyleMatcher;
import org.zkoss.zss.model.util.FontMatcher;
import org.zkoss.zss.range.SRanges;
import org.zkoss.zss.range.impl.DataValidationHelper;

public class ValidationTest {

	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	protected SSheet initialDataGrid(SSheet sheet){
		return sheet;
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
		dv1.setFormula1("A1:C1");
		
		Assert.assertEquals(3, dv1.getNumOfValue1());
		Assert.assertEquals(0, dv1.getNumOfValue2());
		Assert.assertEquals(1D, dv1.getValue1(0));
		Assert.assertEquals(2D, dv1.getValue1(1));
		Assert.assertEquals(3D, dv1.getValue1(2));
		
		
		dv2.setValidationType(ValidationType.INTEGER);
		((AbstractDataValidationAdv)dv2).setFormulas("A1", "C1");
		Assert.assertEquals(1, dv2.getNumOfValue1());
		Assert.assertEquals(1, dv2.getNumOfValue2());
		Assert.assertEquals(1D, dv2.getValue1(0));
		Assert.assertEquals(3D, dv2.getValue2(0));
		
		dv3.setValidationType(ValidationType.INTEGER);
		dv3.setFormula1("AVERAGE(A1:C1)");
		dv3.setFormula2("SUM(A1:C1)");
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
	
	@Test
	public void testDataValidationHelperInteger(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		sheet1.getCell(0, 0).setValue(1D);//min
		sheet1.getCell(0, 1).setValue(3D);//max
		SRanges.range(sheet1,0,2).setEditText("2013/1/1");//day start
		SRanges.range(sheet1,0,3).setEditText("2013/2/1");//day end
		SRanges.range(sheet1,0,4).setEditText("12:00");//time start
		SRanges.range(sheet1,0,5).setEditText("14:00");//time end
		
		Assert.assertEquals(CellType.NUMBER, sheet1.getCell(0, 2).getType());
		Assert.assertEquals("2013/1/1", SRanges.range(sheet1,0,2).getCellFormatText());
		Assert.assertEquals(CellType.NUMBER, sheet1.getCell(0, 3).getType());
		Assert.assertEquals("2013/2/1", SRanges.range(sheet1,0,3).getCellFormatText());
		Assert.assertEquals(CellType.NUMBER, sheet1.getCell(0, 4).getType());
		Assert.assertEquals("12:00", SRanges.range(sheet1,0,4).getCellFormatText());
		Assert.assertEquals(CellType.NUMBER, sheet1.getCell(0, 5).getType());
		Assert.assertEquals("14:00", SRanges.range(sheet1,0,5).getCellFormatText());
		
		
		SDataValidation dv1 = sheet1.addDataValidation(new CellRegion(1,1));
		//test any
		Assert.assertTrue(new DataValidationHelper(dv1).validate("123", "General"));
		
		
		//test integer
		dv1.setValidationType(ValidationType.INTEGER);
		((AbstractDataValidationAdv)dv1).setFormulas("1", "3");
		dv1.setOperatorType(OperatorType.BETWEEN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1.3", "General"));//not integer
		
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("3", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("4", "General"));
		
		dv1.setOperatorType(OperatorType.NOT_BETWEEN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("3", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("4", "General"));
		
		dv1.setOperatorType(OperatorType.EQUAL);
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("3", "General"));
		
		dv1.setOperatorType(OperatorType.NOT_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("3", "General"));
		
		dv1.setOperatorType(OperatorType.GREATER_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		
		dv1.setOperatorType(OperatorType.GREATER_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		
		dv1.setOperatorType(OperatorType.LESS_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("0", "General"));
		
		dv1.setOperatorType(OperatorType.LESS_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("0", "General"));
	}
	
	@Test
	public void testDataValidationHelperDecimal(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		sheet1.getCell(0, 0).setValue(1D);//min
		sheet1.getCell(0, 1).setValue(2D);//max SUM(A1:A2)
		
		
		SDataValidation dv1 = sheet1.addDataValidation(new CellRegion(1,1));
		
		
		//test integer
		dv1.setValidationType(ValidationType.DECIMAL);
		((AbstractDataValidationAdv)dv1).setFormulas("A1", "SUM(A1:B1)");//1-3
		dv1.setOperatorType(OperatorType.BETWEEN);
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1.3", "General"));//not integer
		
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("3", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("4", "General"));
		
		dv1.setOperatorType(OperatorType.NOT_BETWEEN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("3", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("4", "General"));
		
		dv1.setOperatorType(OperatorType.EQUAL);
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("3", "General"));
		
		dv1.setOperatorType(OperatorType.NOT_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("3", "General"));
		
		dv1.setOperatorType(OperatorType.GREATER_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		
		dv1.setOperatorType(OperatorType.GREATER_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("0", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2", "General"));
		
		dv1.setOperatorType(OperatorType.LESS_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("0", "General"));
		
		dv1.setOperatorType(OperatorType.LESS_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("1", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("0", "General"));
	}
	
	@Test
	public void testDataValidationHelperTextlength(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		sheet1.getCell(0, 0).setValue(1D);//min
		sheet1.getCell(0, 1).setValue(2D);//max SUM(A1:A2)
		
		
		SDataValidation dv1 = sheet1.addDataValidation(new CellRegion(1,1));
		
		
		//test integer
		dv1.setValidationType(ValidationType.TEXT_LENGTH);
		((AbstractDataValidationAdv)dv1).setFormulas("A1", "SUM(A1:B1)");//1-3
		dv1.setOperatorType(OperatorType.BETWEEN);
		
		Assert.assertTrue(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("AA", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("AAA", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("AAAA", "General"));
		
		dv1.setOperatorType(OperatorType.NOT_BETWEEN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("AA", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("AAA", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("AAAA", "General"));
		
		dv1.setOperatorType(OperatorType.EQUAL);
		Assert.assertTrue(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("AA", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("AAA", "General"));
		
		dv1.setOperatorType(OperatorType.NOT_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("AA", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("AAA", "General"));
		
		dv1.setOperatorType(OperatorType.GREATER_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("AA", "General"));
		
		dv1.setOperatorType(OperatorType.GREATER_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("AA", "General"));
		
		dv1.setOperatorType(OperatorType.LESS_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("AA", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("", "General"));
		
		dv1.setOperatorType(OperatorType.LESS_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("AA", "General"));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("A", "General"));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("", "General"));
	}
	
	@Test
	public void testDataValidationHelperDate(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		SRanges.range(sheet1,0,0).setEditText("2013/1/10");//day start
		SRanges.range(sheet1,0,1).setEditText("2013/2/1");//day end
		SRanges.range(sheet1,0,2).setEditText("12:00");//time start
		SRanges.range(sheet1,0,3).setEditText("14:00");//time end
		
		
		SDataValidation dv1 = sheet1.addDataValidation(new CellRegion(1,1));
		
		
		//test integer
		dv1.setValidationType(ValidationType.DATE);
		((AbstractDataValidationAdv)dv1).setFormulas("A1", "B1");//2013/1/10 - 2013/2/1
		dv1.setOperatorType(OperatorType.BETWEEN);
		String format = "yyyy/m/d";
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/15", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/2/1", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/9", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/2/2", format));
		
		dv1.setOperatorType(OperatorType.NOT_BETWEEN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/15", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/2/1", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/9", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/2/2", format));
		
		dv1.setOperatorType(OperatorType.EQUAL);
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/15", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/2/1", format));
		
		dv1.setOperatorType(OperatorType.NOT_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/15", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/2/1", format));
		
		dv1.setOperatorType(OperatorType.GREATER_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/9", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/15", format));
		
		dv1.setOperatorType(OperatorType.GREATER_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/9", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/15", format));
		
		dv1.setOperatorType(OperatorType.LESS_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/15", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/9", format));
		
		dv1.setOperatorType(OperatorType.LESS_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/15", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("2013/1/10", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("2013/1/9", format));
	}
	
	@Test
	public void testDataValidationHelperTime(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		SRanges.range(sheet1,0,0).setEditText("12:00");//time start
		SRanges.range(sheet1,0,1).setEditText("14:00");//time end
		
		
		SDataValidation dv1 = sheet1.addDataValidation(new CellRegion(1,1));
		
		
		//test integer
		dv1.setValidationType(ValidationType.TIME);
		((AbstractDataValidationAdv)dv1).setFormulas("A1", "B1");//12:00 - 14:00
		dv1.setOperatorType(OperatorType.BETWEEN);
		String format = "h:mm";
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:01", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("14:00", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("11:59", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("14:01", format));
		
		dv1.setOperatorType(OperatorType.NOT_BETWEEN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:01", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("14:00", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("11:59", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("14:01", format));
		
		dv1.setOperatorType(OperatorType.EQUAL);
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:01", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("14:00", format));
		
		dv1.setOperatorType(OperatorType.NOT_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:01", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("14:00", format));
		
		dv1.setOperatorType(OperatorType.GREATER_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("11:59", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:01", format));
		
		dv1.setOperatorType(OperatorType.GREATER_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("11:59", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:01", format));
		
		dv1.setOperatorType(OperatorType.LESS_OR_EQUAL);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:01", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("11:59", format));
		
		dv1.setOperatorType(OperatorType.LESS_THAN);
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:01", format));
		Assert.assertFalse(new DataValidationHelper(dv1).validate("12:00", format));
		Assert.assertTrue(new DataValidationHelper(dv1).validate("11:59", format));
	}
	
	
	@Test
	public void testDataValidationHelperList(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		SRanges.range(sheet1,0,0).setEditText("A");
		SRanges.range(sheet1,0,1).setEditText("B");
		SRanges.range(sheet1,0,2).setEditText("C");
		SRanges.range(sheet1,1,0).setEditText("1");
		SRanges.range(sheet1,1,1).setEditText("2");
		SRanges.range(sheet1,1,2).setEditText("3");
		SRanges.range(sheet1,2,0).setEditText("2013/1/1");
		SRanges.range(sheet1,2,1).setEditText("2013/1/2");
		SRanges.range(sheet1,2,2).setEditText("2013/1/3");
		
		SDataValidation dv0 = sheet1.addDataValidation(new CellRegion(0,3));
		
		String dateFormat = "yyyy/m/d";
		String numberFormat = "General";
		//test integer
		dv0.setValidationType(ValidationType.LIST);
		dv0.setFormula1("{1,2,3}");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("1", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("2", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("3", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("0", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("4", numberFormat));
		
		dv0.setFormula1("{\"A\",\"B\",\"C\"}");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("A", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("B", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("C", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("D", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("E", numberFormat));
		
		dv0.setFormula1("A1:C1");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("A", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("B", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("C", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("D", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("E", numberFormat));
		
		dv0.setFormula1("A2:C2");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("1", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("2", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("3", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("0", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("4", numberFormat));
		
		dv0.setFormula1("A3:C3");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("2013/1/1", dateFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("2013/1/2", dateFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("2013/1/3", dateFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("2013/1/4", dateFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("2013/1/5", dateFormat));
	}
	@Test
	public void testDataValidationHelperFormula(){
		SBook book = SBooks.createBook("book1");
		SSheet sheet1 = initialDataGrid(book.createSheet("Sheet1"));
		SRanges.range(sheet1,0,0).setEditText("A");
		SRanges.range(sheet1,0,1).setEditText("B");
		SRanges.range(sheet1,0,2).setEditText("C");
		SRanges.range(sheet1,0,3).setEditText("D");
		SRanges.range(sheet1,1,0).setEditText("1");
		SRanges.range(sheet1,1,1).setEditText("2");
		SRanges.range(sheet1,1,2).setEditText("3");
		SRanges.range(sheet1,1,3).setEditText("4");
		SRanges.range(sheet1,2,0).setEditText("2013/1/1");
		SRanges.range(sheet1,2,1).setEditText("2013/1/2");
		SRanges.range(sheet1,2,2).setEditText("2013/1/3");
		SRanges.range(sheet1,2,3).setEditText("2013/1/4");
		
		SDataValidation dv0 = sheet1.addDataValidation(new CellRegion(0,4));
		
		String dateFormat = "yyyy/m/d";
		String numberFormat = "General";
		//test integer
		dv0.setValidationType(ValidationType.LIST);
		dv0.setFormula1("{1,2,3}");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=A2", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=B2", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=C2", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=D2", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=E2", numberFormat));
		
		dv0.setFormula1("{\"A\",\"B\",\"C\"}");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=A1", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=B1", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=C1", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=D1", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=E1", numberFormat));
		
		dv0.setFormula1("A1:C1");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=A1", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=B1", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=C1", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=D1", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=E1", numberFormat));
		
		dv0.setFormula1("A2:C2");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=A2", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=B2", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=C2", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=D2", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=E2", numberFormat));
		
		dv0.setFormula1("A3:C3");
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=A3", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=B3", numberFormat));
		Assert.assertTrue(new DataValidationHelper(dv0).validate("=C3", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=D3", numberFormat));
		Assert.assertFalse(new DataValidationHelper(dv0).validate("=E3", numberFormat));
	}
}
