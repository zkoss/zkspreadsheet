package org.zkoss.zss.model;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.ArrayList;
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
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.InvalidateModelOpException;
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
import org.zkoss.zss.model.impl.sys.DependencyTableImpl;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.util.CellStyleMatcher;
import org.zkoss.zss.model.util.FontMatcher;
import org.zkoss.zss.model.util.Validations;
import org.zkoss.zss.range.SRanges;

public class IssueTest {

	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
		SheetImpl.DEBUG = true;
	}
	
	@Test 
	public void testZSS581(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		SSheet sheet3 = book.createSheet("Sheet3");
		
		sheet1.getCell("A1").setValue("=Sheet2!A1");
		sheet2.getCell("A1").setValue("=Sheet3!A1");
		sheet3.getCell("A1").setValue("3");
		
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
//		System.out.println(">>>>>>>>case1");
//		((DependencyTableImpl)table).dump();
		
		SRanges.range(sheet2).setSheetName("Two");
//		System.out.println(">>>>>>>>case2");
//		((DependencyTableImpl)table).dump();
		Assert.assertEquals("Two!A1",sheet1.getCell("A1").getFormulaValue());
		
		SRanges.range(sheet3).setSheetName("Three");
//		System.out.println(">>>>>>>>case3");
//		((DependencyTableImpl)table).dump();
		Assert.assertEquals("Three!A1",sheet2.getCell("A1").getFormulaValue());
		
			
		sheet3.getCell("A1").setValue(10);
		Assert.assertEquals(10D, sheet1.getCell("A1").getValue());
		Assert.assertEquals(10D, sheet2.getCell("A1").getValue());
		
	}
	
	@Test
	public void testZSS581_furthercheck(){
		SBook book = SBooks.createBook("book1");
		book.getBookSeries().setAutoFormulaCacheClean(true);
		
		SSheet sheet1 = book.createSheet("Sheet1");
		SSheet sheet2 = book.createSheet("Sheet2");
		SSheet sheet3 = book.createSheet("Sheet3");
		
		sheet1.getCell("A1").setValue("=Sheet2!A1");
		sheet2.getCell("A1").setValue("=Sheet3!A1");
		sheet3.getCell("A1").setValue("3");
		
		DependencyTable table = ((AbstractBookSeriesAdv)book.getBookSeries()).getDependencyTable();
		((DependencyTableImpl)table).dump();
		
		List<Ref> refs = new ArrayList<Ref>(table.getDirectDependents(new RefImpl(book.getBookName())));
		Assert.assertEquals(2, refs.size());
		Ref ref = refs.get(0);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());		
		
		refs = new ArrayList<Ref>(table.getDirectDependents(new RefImpl(book.getBookName(),sheet3.getSheetName())));
		Assert.assertEquals(1, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		refs = new ArrayList<Ref>(table.getDirectDependents(new RefImpl(book.getBookName(),sheet2.getSheetName())));
		Assert.assertEquals(2, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		refs = new ArrayList<Ref>(table.getDependents(new RefImpl(book.getBookName(),sheet3.getSheetName())));
		Assert.assertEquals(2, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		
		refs = new ArrayList<Ref>(table.getDependents(new RefImpl(book.getBookName(),sheet2.getSheetName())));
		Assert.assertEquals(2, refs.size());
		ref = refs.get(0);
		Assert.assertEquals("Sheet1", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());
		ref = refs.get(1);
		Assert.assertEquals("Sheet2", ref.getSheetName());
		Assert.assertEquals(0, ref.getRow());
		Assert.assertEquals(0, ref.getColumn());
		Assert.assertEquals(0, ref.getLastRow());
		Assert.assertEquals(0, ref.getLastColumn());		
	}	
	
}
