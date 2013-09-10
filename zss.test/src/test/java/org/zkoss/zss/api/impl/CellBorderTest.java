package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.zss.SpreadsheetAgent;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.ui.Spreadsheet;



/**
 * Test case for the function "display Excel files" for 2007 and 2003
 * Testing for the sheet "cell-border".
 * 
 * Because the specification of color is unclear, we don't test color-related feature.
 * 
 * @author Hawk
 *
 */
public class CellBorderTest extends DisplayExcelTest{
	
	public CellBorderTest(){
		this("/display.zul");
	}
	
	protected CellBorderTest(String testPage){
		super(testPage);
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("cell-border");
		sheet = zss.as(Spreadsheet.class).getBook().getSheet("cell-border");
	}


	@Test
	public void testBorderPosition(){
		//bottom
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 2, 1).getBorderBottom());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 2, 1).getBorderTop());

		//top
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 2, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 2, 2).getBorderRight());
		
		//left
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 2, 3).getBorderLeft());
		
		//right
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 2, 4).getBorderRight());
	}
	
	@Test
	public void testBorderStyle(){
		//only styles that can be rendered 
		assertEquals(CellStyle.BorderType.HAIR,getCellStyle(sheet, 4, 1).getBorderBottom());
		assertEquals(CellStyle.BorderType.DOTTED,getCellStyle(sheet, 4, 2).getBorderBottom());
		assertEquals(CellStyle.BorderType.DASHED,getCellStyle(sheet, 4, 3).getBorderBottom());
		
	}
	
	@Test
	public void testAllBorders4Cells(){
		for (int row = 8 ; row<=9 ; row++){
			for (int column=1 ; column <=2 ; column++){
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderTop());
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderLeft());
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderRight());
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderBottom());
			}
		}
	}
	
	@Test
	public void testOutlineBorders4Cells(){
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 8, 4).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 8, 4).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 8, 4).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 8, 4).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 8, 5).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 8, 5).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 8, 5).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 8, 5).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 9, 4).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 9, 4).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 9, 4).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 9, 4).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 9, 5).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 9, 5).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 9, 5).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 9, 5).getBorderBottom());;
	}
	
	@Test
	public void testInsideBorders4Cells(){
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 11, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 11, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 11, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 11, 1).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 11, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 11, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 11, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 11, 2).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 12, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 12, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 12, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 12, 1).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 12, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 12, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 12, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 12, 2).getBorderBottom());;
	}
	
	@Test
	public void testOutlineBorders9Cells(){

		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 1).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 2).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 3).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 3).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 3).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 3).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 1).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 2).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 3).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 3).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 3).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 3).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 1).getBorderRight());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 2).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 3).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 3).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 3).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 3).getBorderBottom());
		
	}
	
	@Test
	public void testInsideBorders9Cells(){
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 5).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 5).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 5).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 5).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 6).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 6).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 6).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 6).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 7).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 7).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 18, 7).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 18, 7).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 5).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 5).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 5).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 5).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 6).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 6).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 6).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 6).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 7).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 7).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 19, 7).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 19, 7).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 5).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 5).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 5).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 5).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 6).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 6).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 6).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 6).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 7).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 20, 7).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 7).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 20, 7).getBorderBottom());
	}
	
	@Test
	public void testHorizontalBorders9Cells(){
		for (int row=22; row<=24 ; row++){
			for (int column = 1 ; column<=3 ; column++){
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderTop());
				assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, row, column).getBorderLeft());
				assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, row, column).getBorderRight());
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderBottom());
			}
		}
	}
	
	@Test
	public void testVerticalBorders9Cells(){
		for (int row=22; row<=24 ; row++){
			for (int column = 5 ; column<=7 ; column++){
				assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, row, column).getBorderTop());
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderLeft());
				assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, row, column).getBorderRight());
				assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, row, column).getBorderBottom());
			}
		}
	}
	
	@Test
	public void testMergedCellBorder(){
		//merge 2 cells
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 27, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 1).getBorderBottom());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 27, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 2).getBorderBottom());

		//merge 3 cells
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 4).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 4).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 27, 4).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 4).getBorderBottom());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 5).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 27, 5).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 27, 5).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 5).getBorderBottom());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 6).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 27, 6).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 6).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 27, 6).getBorderBottom());
		
	}
	
	@Test
	public void testMergedVerticalCellBorder(){
		//merge cells vertically
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 30, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 30, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 30, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 30, 1).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 31, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 31, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 1).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 32, 1).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 32, 1).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 32, 1).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 32, 1).getBorderBottom());

	}
	
	@Test
	public void testMergedVerticalHorizontalCellBorder(){
		//merge cells vertically and horizontally
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 30, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 30, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 30, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 30, 2).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 30, 3).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 30, 3).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 30, 3).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 30, 3).getBorderBottom());
		
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 31, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 2).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 3).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 3).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 31, 3).getBorderRight());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 31, 3).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 32, 2).getBorderTop());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 32, 2).getBorderLeft());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 32, 2).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 32, 2).getBorderBottom());
		
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 32, 3).getBorderTop());
		assertEquals(CellStyle.BorderType.NONE,getCellStyle(sheet, 32, 3).getBorderLeft());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 32, 3).getBorderRight());
		assertEquals(CellStyle.BorderType.THIN,getCellStyle(sheet, 32, 3).getBorderBottom());

	}
}
