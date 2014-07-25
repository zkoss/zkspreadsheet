package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.AssertUtil;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Sheet;

/**
 * @author dennis
 *
 */
public class BorderTest {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
	
	@Test
	public void testBorder() throws IOException {
		testBorder(Util.loadBook("bordercolor.xlsx"));
		testBorder(Util.loadBook("bordercolor.xls"));
		testMakeBorder(Util.loadBook("blank.xlsx"));
		testMakeBorder(Util.loadBook("blank.xls"));
	}
	
	public void testMakeBorder(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B1"),ApplyBorderType.EDGE_BOTTOM,BorderType.THIN,"#ff0000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"C1"),ApplyBorderType.EDGE_BOTTOM,BorderType.THIN,"#00ff00");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"D1"),ApplyBorderType.EDGE_BOTTOM,BorderType.THIN,"#0000ff");
		
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B3"),ApplyBorderType.EDGE_BOTTOM,BorderType.THIN,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"C3"),ApplyBorderType.EDGE_TOP,BorderType.THIN,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"D3"),ApplyBorderType.EDGE_LEFT,BorderType.THIN,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"E3"),ApplyBorderType.EDGE_RIGHT,BorderType.THIN,"#000000");
		
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B5"),ApplyBorderType.EDGE_BOTTOM,BorderType.HAIR,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"C5"),ApplyBorderType.EDGE_BOTTOM,BorderType.DOTTED,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"D5"),ApplyBorderType.EDGE_BOTTOM,BorderType.DASHED,"#000000");
		
		
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B9:C10"),ApplyBorderType.FULL,BorderType.THIN,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"E9:F10"),ApplyBorderType.OUTLINE,BorderType.THIN,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B12:C13"),ApplyBorderType.INSIDE,BorderType.THIN,"#000000");
		
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B15:C16"),ApplyBorderType.OUTLINE,BorderType.THIN,"#0000ff");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B15:C16"),ApplyBorderType.INSIDE,BorderType.THIN,"#ff0000");
		
		
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B19:D21"),ApplyBorderType.OUTLINE,BorderType.THIN,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"F19:H21"),ApplyBorderType.INSIDE,BorderType.THIN,"#000000");
		
		CellOperationUtil.applyBorder(Ranges.range(sheet,"B22:D26"),ApplyBorderType.INSIDE_HORIZONTAL,BorderType.THIN,"#000000");
		CellOperationUtil.applyBorder(Ranges.range(sheet,"E23:I25"),ApplyBorderType.INSIDE_VERTICAL,BorderType.THIN,"#000000");
		
		testBorder(book);
		
		book = Util.swap(book);
		
		testBorder(book);
	}
	public void testBorder(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		
		Range r = Ranges.range(sheet,"B1");
		
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#ff0000");
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);
		
		
		r = Ranges.range(sheet,"C1");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#00ff00");
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);
		
		r = Ranges.range(sheet,"D1");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#0000ff");
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);
		
		
		r = Ranges.range(sheet,"B3");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);
		
		r = Ranges.range(sheet,"C3");
		AssertUtil.assertTopBorder(r, BorderType.THIN, "#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE,null);
		AssertUtil.assertRightBorder(r, BorderType.THIN, "#000000");
		
		r = Ranges.range(sheet,"D3");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE,null);
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);
		
		r = Ranges.range(sheet,"E3");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE,null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		
		r = Ranges.range(sheet,"B5");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.HAIR,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);

		r = Ranges.range(sheet,"C5");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.DOTTED,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);
		
		r = Ranges.range(sheet,"D5");
		AssertUtil.assertTopBorder(r, BorderType.NONE,null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE,null);
		AssertUtil.assertBottomBorder(r, BorderType.DASHED,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE,null);
		
		
		//all 
		r = Ranges.range(sheet,"B9");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"C9");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"B10");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"C10");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		
		//outline
		r = Ranges.range(sheet,"E9");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"F9");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"E10");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"F10");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		
		//inside
		r = Ranges.range(sheet,"B12");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"C12");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"B13");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"C13");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);		
		
		//blue/red//all 
		r = Ranges.range(sheet,"B15");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#0000ff");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#0000ff");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#ff0000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#ff0000");
		r = Ranges.range(sheet,"C15");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#0000ff");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#ff0000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#ff0000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#0000ff");
		r = Ranges.range(sheet,"B16");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#ff0000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#0000ff");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#0000ff");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#ff0000");
		r = Ranges.range(sheet,"C16");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#ff0000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#ff0000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#0000ff");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#0000ff");
		
		//9 cells
		r = Ranges.range(sheet,"B19");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"C19");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"D19");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"B20");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"C20");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"D20");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"B21");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"C21");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"D21");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		
		//9 inside
		r = Ranges.range(sheet,"F19");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"G19");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"H19");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"F20");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"G20");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"H20");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"F21");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"G21");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"H21");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		
		//9 h
		r = Ranges.range(sheet,"B23");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"C23");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"D23");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"B24");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"C24");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"D24");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"B25");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"C25");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"D25");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"B26");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"C26");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"D26");
		AssertUtil.assertTopBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		
		//9 v
		r = Ranges.range(sheet,"E23");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"F23");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"G23");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r,BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"H23");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"I23");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"E24");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"F24");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"G24");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"H24");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"I24");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
		r = Ranges.range(sheet,"E25");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.NONE, null);
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"F25");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"G25");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"H25");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r,BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.THIN,"#000000");
		r = Ranges.range(sheet,"I25");
		AssertUtil.assertTopBorder(r, BorderType.NONE, null);
		AssertUtil.assertLeftBorder(r, BorderType.THIN,"#000000");
		AssertUtil.assertBottomBorder(r, BorderType.NONE, null);
		AssertUtil.assertRightBorder(r, BorderType.NONE, null);
	}
	
	
}
