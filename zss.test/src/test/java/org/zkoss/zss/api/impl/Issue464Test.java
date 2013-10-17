package org.zkoss.zss.api.impl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.AssertUtil;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Sheet;

/**
 * Complex tests for issue 464
 * @author dennis
 *
 */
public class Issue464Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}
	
	@Test
	public void testApplyBorder() throws IOException {
		testApplyBorder(Util.loadBook(this,"book/blank.xls"));
		testApplyBorder(Util.loadBook(this,"book/blank.xlsx"));
		
	}
	
	private void testApplyBorder(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyBorder(r,ApplyBorderType.FULL, BorderType.MEDIUM, "#ff0000");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderBottom());
				Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderTop());
				Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderLeft());
				Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderRight());
				Assert.assertEquals(location, "#ff0000",style.getBorderBottomColor().getHtmlColor());
				Assert.assertEquals(location, "#ff0000",style.getBorderTopColor().getHtmlColor());
				Assert.assertEquals(location, "#ff0000",style.getBorderLeftColor().getHtmlColor());
				Assert.assertEquals(location, "#ff0000",style.getBorderRightColor().getHtmlColor());
			}
		}
		
		r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyBorder(r,ApplyBorderType.FULL, BorderType.NONE, "#000000");//color assign doesn't take effect when assign none
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
				Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
				Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
				Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
			}
		}
		
		r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyBorder(r,ApplyBorderType.OUTLINE, BorderType.MEDIUM, "#ffff00");//color assign doesn't take effect when assign none
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i==0){
					if(j==0){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderTop());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderTopColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderLeftColor().getHtmlColor());
					}else if(j==column){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderRight()); 

						Assert.assertEquals(location, "#ffff00",style.getBorderTopColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderRightColor().getHtmlColor()); 
					}else{
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
					}
				}else if(i==row){
					if(j==0){
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderBottomColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderLeftColor().getHtmlColor());
					}else if(j==column){
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderBottomColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderRightColor().getHtmlColor());
					}else{
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
					}
				}else{
					if(j==0){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderLeftColor().getHtmlColor());
					}else if(j==column){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderRightColor().getHtmlColor());
					}else{
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
					}
				}
				
			}
		}
		
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i==0){
					if(j==0){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderTop());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderTopColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderLeftColor().getHtmlColor());
					}else if(j==column){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderRight()); 

						Assert.assertEquals(location, "#ffff00",style.getBorderTopColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderRightColor().getHtmlColor()); 
					}else{
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
					}
				}else if(i==row){
					if(j==0){
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderBottomColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderLeftColor().getHtmlColor());
					}else if(j==column){
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderBottomColor().getHtmlColor());
						Assert.assertEquals(location, "#ffff00",style.getBorderRightColor().getHtmlColor());
					}else{
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
					}
				}else{
					if(j==0){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderLeftColor().getHtmlColor());
					}else if(j==column){
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.MEDIUM,style.getBorderRight());
						
						Assert.assertEquals(location, "#ffff00",style.getBorderRightColor().getHtmlColor());
					}else{
						Assert.assertEquals(location, BorderType.NONE,style.getBorderBottom());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderTop());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderLeft());
						Assert.assertEquals(location, BorderType.NONE,style.getBorderRight());
					}
				}
				
			}
		}
	}
	//CELL STYLE
	@Test
	public void testApplyBackgroundColor() throws IOException {
		testApplyBackgroundColor(Util.loadBook(this,"book/blank.xls"));
		testApplyBackgroundColor(Util.loadBook(this,"book/blank.xlsx"));
	}
	
	private void testApplyBackgroundColor(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyBackgroundColor(r,"#ff0000");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, "#ff0000", style.getBackgroundColor().getHtmlColor());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyBackgroundColor(r,"#ffff00");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "#ffff00", style.getBackgroundColor().getHtmlColor());
				}else{
					Assert.assertEquals(location, "#ff0000", style.getBackgroundColor().getHtmlColor());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "#ffff00", style.getBackgroundColor().getHtmlColor());
				}else{
					Assert.assertEquals(location, "#ff0000", style.getBackgroundColor().getHtmlColor());
				}
			}
		}
	}
	
	@Test
	public void testApplyAlignment() throws IOException {
		testApplyAlignment(Util.loadBook(this,"book/blank.xls"));
		testApplyAlignment(Util.loadBook(this,"book/blank.xlsx"));
	}
	
	private void testApplyAlignment(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyAlignment(r, Alignment.CENTER);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, Alignment.CENTER, style.getAlignment());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyAlignment(r, Alignment.RIGHT);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, Alignment.RIGHT, style.getAlignment());
				}else{
					Assert.assertEquals(location, Alignment.CENTER, style.getAlignment());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, Alignment.RIGHT, style.getAlignment());
				}else{
					Assert.assertEquals(location, Alignment.CENTER, style.getAlignment());
				}
			}
		}
	}
	
	
	@Test
	public void testApplyVerticalAlignment() throws IOException {
		testApplyVerticalAlignment(Util.loadBook(this,"book/blank.xls"));
		testApplyVerticalAlignment(Util.loadBook(this,"book/blank.xlsx"));
	}
	
	private void testApplyVerticalAlignment(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyVerticalAlignment(r, VerticalAlignment.CENTER);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, VerticalAlignment.CENTER, style.getVerticalAlignment());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyVerticalAlignment(r, VerticalAlignment.BOTTOM);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, VerticalAlignment.BOTTOM, style.getVerticalAlignment());
				}else{
					Assert.assertEquals(location, VerticalAlignment.CENTER, style.getVerticalAlignment());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, VerticalAlignment.BOTTOM, style.getVerticalAlignment());
				}else{
					Assert.assertEquals(location, VerticalAlignment.CENTER, style.getVerticalAlignment());
				}
			}
		}
	}
	
	
	@Test
	public void testApplyFormat() throws IOException {
		testApplyFormat(Util.loadBook(this,"book/blank.xls"));
		testApplyFormat(Util.loadBook(this,"book/blank.xlsx"));
	}
	
	private void testApplyFormat(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyDataFormat(r, "yyyymmdd");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, "yyyymmdd", style.getDataFormat());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyDataFormat(r, "#.0");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "#.0", style.getDataFormat());
				}else{
					Assert.assertEquals(location, "yyyymmdd", style.getDataFormat());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "#.0", style.getDataFormat());
				}else{
					Assert.assertEquals(location, "yyyymmdd", style.getDataFormat());
				}
			}
		}
	}
	
	@Test
	public void testApplyTextwrap() throws IOException {
		testApplyTextwrap(Util.loadBook(this,"book/blank.xls"));
		testApplyTextwrap(Util.loadBook(this,"book/blank.xlsx"));	
	}
	private void testApplyTextwrap(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyWrapText(r, true);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, true, style.isWrapText());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyWrapText(r, false);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, false, style.isWrapText());
				}else{
					Assert.assertEquals(location, true, style.isWrapText());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, false, style.isWrapText());
				}else{
					Assert.assertEquals(location, true, style.isWrapText());
				}
			}
		}
	}
	
	//FONT
	
	@Test
	public void testApplyFontName() throws IOException {
		testApplyFontName(Util.loadBook(this,"book/blank.xls"));
		testApplyFontName(Util.loadBook(this,"book/blank.xlsx"));
		
	}
	
	private void testApplyFontName(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyFontName(r, "Georgia");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, "Georgia", style.getFont().getFontName());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyFontName(r, "Tahoma");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "Tahoma", style.getFont().getFontName());
				}else{
					Assert.assertEquals(location, "Georgia", style.getFont().getFontName());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "Tahoma", style.getFont().getFontName());
				}else{
					Assert.assertEquals(location, "Georgia", style.getFont().getFontName());
				}
			}
		}
	}
	
	@Test
	public void testApplyFontSize() throws IOException {
		testApplyFontSize(Util.loadBook(this,"book/blank.xls"));
		testApplyFontSize(Util.loadBook(this,"book/blank.xlsx"));	
	}
	private void testApplyFontSize(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyFontHeight(r, 20);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, 20, style.getFont().getFontHeight());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyFontHeight(r, 40);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, 40, style.getFont().getFontHeight());
				}else{
					Assert.assertEquals(location, 20, style.getFont().getFontHeight());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, 40, style.getFont().getFontHeight());
				}else{
					Assert.assertEquals(location, 20, style.getFont().getFontHeight());
				}
			}
		}
	}
	
	
	@Test
	public void testApplyFontColor() throws IOException {
		testApplyFontColor(Util.loadBook(this,"book/blank.xls"));
		testApplyFontColor(Util.loadBook(this,"book/blank.xlsx"));	
	}
	private void testApplyFontColor(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyFontColor(r, "#0000ff");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, "#0000ff", style.getFont().getColor().getHtmlColor());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyFontColor(r, "#ff00ff");
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "#ff00ff", style.getFont().getColor().getHtmlColor());
				}else{
					Assert.assertEquals(location, "#0000ff", style.getFont().getColor().getHtmlColor());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, "#ff00ff", style.getFont().getColor().getHtmlColor());
				}else{
					Assert.assertEquals(location, "#0000ff", style.getFont().getColor().getHtmlColor());
				}
			}
		}
	}
	
	
	@Test
	public void testApplyFontBold() throws IOException {
		testApplyFontBold(Util.loadBook(this,"book/blank.xls"));
		testApplyFontBold(Util.loadBook(this,"book/blank.xlsx"));	
	}
	private void testApplyFontBold(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyFontBoldweight(r, Boldweight.BOLD);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, Boldweight.BOLD, style.getFont().getBoldweight());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyFontBoldweight(r, Boldweight.NORMAL);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, Boldweight.NORMAL, style.getFont().getBoldweight());
				}else{
					Assert.assertEquals(location, Boldweight.BOLD, style.getFont().getBoldweight());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, Boldweight.NORMAL, style.getFont().getBoldweight());
				}else{
					Assert.assertEquals(location, Boldweight.BOLD, style.getFont().getBoldweight());
				}
			}
		}
	}
	
	
	@Test
	public void testApplyFontItalic() throws IOException {
		testApplyFontItalic(Util.loadBook(this,"book/blank.xls"));
		testApplyFontItalic(Util.loadBook(this,"book/blank.xlsx"));	
	}
	private void testApplyFontItalic(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyFontItalic(r, true);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, true, style.getFont().isItalic());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyFontItalic(r, false);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, false, style.getFont().isItalic());
				}else{
					Assert.assertEquals(location, true, style.getFont().isItalic());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, false, style.getFont().isItalic());
				}else{
					Assert.assertEquals(location, true, style.getFont().isItalic());
				}
			}
		}
	}
	
	
	@Test
	public void testApplyFontStrike() throws IOException {
		testApplyFontStrike(Util.loadBook(this,"book/blank.xls"));
		testApplyFontStrike(Util.loadBook(this,"book/blank.xlsx"));	
	}
	private void testApplyFontStrike(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyFontStrikeout(r, true);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, true, style.getFont().isStrikeout());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyFontStrikeout(r, false);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, false, style.getFont().isStrikeout());
				}else{
					Assert.assertEquals(location, true, style.getFont().isStrikeout());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, false, style.getFont().isStrikeout());
				}else{
					Assert.assertEquals(location, true, style.getFont().isStrikeout());
				}
			}
		}
	}
	
	
	@Test
	public void testApplyFontUnderline() throws IOException {
		testApplyFontUnderline(Util.loadBook(this,"book/blank.xls"));
		testApplyFontUnderline(Util.loadBook(this,"book/blank.xlsx"));	
	}
	private void testApplyFontUnderline(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		int row = 100;
		int column = 50;
		System.out.println("cell style number before "+book.getPoiBook().getNumCellStyles());
		Range r = Ranges.range(sheet,0,0,row,column);
		CellOperationUtil.applyFontUnderline(r, Underline.SINGLE);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +" Row:"+i+", Column:"+j;
				Assert.assertEquals(location, Underline.SINGLE, style.getFont().getUnderline());
			}
		}
		r = Ranges.range(sheet,1,1,row-1,column-1);
		CellOperationUtil.applyFontUnderline(r, Underline.NONE);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, Underline.NONE, style.getFont().getUnderline());
				}else{
					Assert.assertEquals(location, Underline.SINGLE, style.getFont().getUnderline());
				}
			}
		}
		System.out.println("cell style number after "+book.getPoiBook().getNumCellStyles());
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		for(int i=0;i<=row;i++){
			for(int j=0;j<=column;j++){
				r = Ranges.range(sheet,i,j);
				CellStyle style = r.getCellStyle();
				String location = "At "+Ranges.getCellRefString(i, j) +"(Row:"+i+", Column:"+j+")";
				if(i>0 && j>0 && i< row && j<column){
					Assert.assertEquals(location, Underline.NONE, style.getFont().getUnderline());
				}else{
					Assert.assertEquals(location, Underline.SINGLE, style.getFont().getUnderline());
				}
			}
		}
	}
}
