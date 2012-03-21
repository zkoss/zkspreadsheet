/* ZSSAppTest.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 11:08:45 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import java.util.Iterator;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.Rule;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriver.Navigation;
import org.zkoss.test.Border;
import org.zkoss.test.Browser;
import org.zkoss.test.Color;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.TestingEnvironment;
import org.zkoss.test.zss.Cell.CellType;

import com.google.common.collect.ImmutableList;
import com.google.guiceberry.junit4.GuiceBerryRule;
import com.google.inject.Inject;
import com.google.inject.Injector;
import com.google.inject.assistedinject.FactoryModuleBuilder;
import com.google.inject.name.Names;

/**
 * @author sam
 *
 */
public class ZSSAppTest {

	public static class ZSSAppEnvironment extends TestingEnvironment{
		protected void configure() {
			//ZSSApp URL
			bind(String.class)
			.annotatedWith(Names.named("URL"))
			.toInstance("http://localhost:8088/zssapp/");
			
			bind(String.class)
			.annotatedWith(Names.named("Spreadsheet Id"))
			.toInstance("spreadsheet");
			
			install(new FactoryModuleBuilder().build(Cell.Factory.class));
			install(new FactoryModuleBuilder().build(CellBlock.Factory.class));
			install(new FactoryModuleBuilder().build(CellCache.Factory.class));
			install(new FactoryModuleBuilder().build(CellCacheAggeration.BuilderFactory.class));
			
			super.configure();
		}
	}
	
	protected enum BorderType {
		BOTTOM("borderBottom"),
		TOP("borderTop"),
		LEFT("borderLeft"),
		RIGHT("borderRight"),
		NO("borderNo"),
		ALL("borderAll"),
		OUTSIDE("borderOutside"),
		INSIDE("borderInside"),
		INSIDE_HORIZONTAL("borderInsideHorizontal"),
		INSIDE_VERTICAL("borderInsideVertical");
		
		private String val;
		private BorderType(String val) {
			this.val = val;
		}
		
		@Override
		public String toString() {
			return val;
		}
	}
	
	protected enum AlignType {
		VERTICAL_ALIGN_TOP("verticalAlignTop"),
		VERTICAL_ALIGN_MIDDLE("verticalAlignMiddle"),
		VERTICAL_ALIGN_BOTTOM("verticalAlignBottom"),
		HORIZONTAL_ALIGN_LEFT("horizontalAlignLeft"),
		HORIZONTAL_ALIGN_CENTER("horizontalAlignCenter"),
		HORIZONTAL_ALIGN_RIGHT("horizontalAlignRight");
		
		private String val;
		private AlignType(String val) {
			this.val = val;
		}
		
		@Override
		public String toString() {
			return val;
		}
	}
	
	protected enum PasteSource {
		COPY,
		CUT
	}
	
	protected enum Insert {
		CELL_RIGHT,
		CELL_DOWN
	}
	
	protected enum Delete {
		CELL_LEFT,
		CELL_UP
	}
	
	protected enum FormatNumber {
		ACCOUNTING,
		CURRENCY
	}
	
	@Rule
	public GuiceBerryRule guiceBerry = new GuiceBerryRule(ZSSAppEnvironment.class);
	
	@Inject
	protected Injector injector;
	
	@Inject
	protected WebDriver webDriver;
	
	@Inject
	protected JQueryFactory jqFactory;
	
	protected Navigation navigation;
	protected JavascriptExecutor javascriptExecutor;
	
	@Inject
	protected Spreadsheet spreadsheet;
	
	@Inject
	protected KeyboardDirector keyboardDirector;
	
	@Inject
	protected MouseDirector mouseDirector;
	
	@Inject
	protected ConditionalTimeBlocker timeBlocker;
	
	@Inject
	protected Browser browser;
	
	@Inject
	protected Cell.Factory cellFactory;
	
	@Before
	public void setUp() {
		javascriptExecutor = (JavascriptExecutor) webDriver;
		navigation = webDriver.navigate();
		navigation.refresh();
		
		timeBlocker.waitResponse();
	}
	
	protected boolean isVisible(String stringSelector) {
		JQuery target = jqFactory.create("'" + stringSelector + "'");
		return target.isVisible();
	}
	
	protected void click(String stringSelector) {
		JQuery target = jqFactory.create("'" + stringSelector + "'");
		click(target);
	}
	
	protected void click(JQuery target) {
		mouseDirector.click(target);
	}
	
	protected void focus(int row, int col) {
		spreadsheet.focus(row, col);
	}
	
	protected JQuery jq(String selector) {
		return jqFactory.create("'" + selector + "'");
	}
	
	protected boolean hasBook() {
		return spreadsheet.getSheetCtrl().isExist();
	}
	
	protected Cell getCell(int row, int col) {
		return cellFactory.create(Integer.valueOf(row), Integer.valueOf(col));
	}
	
	protected int getRowfreeze() {
		return spreadsheet.getTopPanel().getRowfreeze();
	}
	
	protected int getColumnfreeze() {
		return spreadsheet.getLeftPanel().getColumnfreeze();
	}
	
	protected CellCache getCellCache(int row, int col) {
		CellCache.Factory factory = injector.getInstance(CellCache.Factory.class);
		return factory.create(row, col);
	}
	
	protected CellType getCellType(int row, int col) {
		return cellFactory.create(row, col).getCellType();
	}
	
	protected CellCacheAggeration.Builder getCellCacheAggerationBuilder(int tRow, int lCol, int bRow, int rCol) {
		CellCacheAggeration.BuilderFactory f = injector.getInstance(CellCacheAggeration.BuilderFactory.class);
		return f.create(tRow, lCol, bRow, rCol);
	}
	
	protected Iterator<Cell> iterator(int tRow, int lCol, int bRow, int rCol) {
		ImmutableList.Builder<Cell> b = new ImmutableList.Builder<Cell>();
		for (int r = tRow; r <= bRow; r++) {
			for (int c = lCol; c <= rCol; c++) {
				b.add(cellFactory.create(r, c));
			}
		}
		return b.build().iterator();
	}
	
	protected void verifyBorder(BorderType type, String borderColor, int tRow, int lCol, int bRow, int rCol, CellCacheAggeration cellCacheAggeration) {
		Iterator<Cell> i = null;
		Border expected = new Border("1px", "solid", borderColor);
		switch (type) {
		case BOTTOM:
			i = iterator(bRow, lCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getBottomBorder());
			}
			break;
		case TOP:
			i = iterator(tRow - 1, lCol, tRow - 1, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getBottomBorder());
			}
			break;
		case LEFT:
			i = iterator(tRow, lCol - 1, bRow, lCol - 1);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getRightBorder());
			}
			break;
		case RIGHT:
			i = iterator(tRow, rCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getRightBorder());
			}
			break;
		case NO:
			i = iterator(tRow, lCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertFalse(c.hasRightBorder());
				Assert.assertFalse(c.hasBottomBorder());
			}
			break;
		case ALL:
			i = iterator(tRow, lCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getRightBorder());
				Assert.assertEquals(expected, c.getBottomBorder());
			}
			break;
		case OUTSIDE:
			//top
			i = iterator(tRow - 1, lCol, tRow - 1, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getBottomBorder());
			}
			
			//right
			i = iterator(tRow, rCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getRightBorder());
			}
			
			//bottom
			i = iterator(bRow, lCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getBottomBorder());
			}
			
			//left
			i = iterator(tRow, lCol - 1, bRow, lCol - 1);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getRightBorder());
			}
			
			//inside cell: remain the same
			i = iterator(tRow, lCol, bRow - 1, rCol - 1);
			while (i.hasNext()) {
				Cell c = i.next();
				int row = c.getRow();
				int col = c.getCol();
				CellCache cc = cellCacheAggeration.getCellCache(row, col);
				Assert.assertEquals(cc.getRightBorder(), c.getRightBorder());
				Assert.assertEquals(cc.getBottomBorder(), c.getBottomBorder());
			}
			break;
		case INSIDE:
			//top: remain the same
			i = iterator(tRow - 1, lCol, tRow - 1, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				int row = c.getRow();
				int col = c.getCol();
				Assert.assertEquals(cellCacheAggeration.getCellCache(row, col).getBottomBorder(), 
						c.getBottomBorder());
			}
			
			//right: remain the same
			i = iterator(tRow, rCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				int row = c.getRow();
				int col = c.getCol();
				Assert.assertEquals(cellCacheAggeration.getCellCache(row, col).getRightBorder()
						, c.getRightBorder());
			}
			
			//bottom: remain the same
			i = iterator(bRow, lCol, bRow, rCol);
			while (i.hasNext()) {
				Cell c = i.next();
				int row = c.getRow();
				int col = c.getCol();
				Assert.assertEquals(cellCacheAggeration.getCellCache(row, col).getBottomBorder(), c.getBottomBorder());
			}
			
			//left: remain the same
			i = iterator(tRow, lCol - 1, bRow, lCol - 1);
			while (i.hasNext()) {
				Cell c = i.next();
				int row = c.getRow();
				int col = c.getCol();
				Assert.assertEquals(cellCacheAggeration.getCellCache(row, col).getRightBorder(), c.getRightBorder());
			}
			
			//inside
			i = iterator(tRow, lCol, bRow - 1, rCol - 1);
			while (i.hasNext()) {
				Cell c = i.next();
				Assert.assertEquals(expected, c.getRightBorder());
				Assert.assertEquals(expected, c.getBottomBorder());
			}
			break;
		case INSIDE_HORIZONTAL:
			i = iterator(tRow, lCol, bRow - 1, rCol - 1);
			while (i.hasNext()) {
				Cell c = i.next();
				int row = c.getRow();
				int col = c.getCol();
				//right border shall remain the same
				Assert.assertEquals(cellCacheAggeration.getCellCache(row, col).getRightBorder(), c.getRightBorder());
				Assert.assertEquals(expected, c.getBottomBorder());
			}
			break;
		case INSIDE_VERTICAL:
			i = iterator(tRow, lCol, bRow - 1, rCol - 1);
			while (i.hasNext()) {
				Cell c = i.next();
				int row = c.getRow();
				int col = c.getCol();
				Assert.assertEquals(expected, c.getRightBorder());
				//bottom border shall remain the same
				Assert.assertEquals(cellCacheAggeration.getCellCache(row, col).getBottomBorder(), c.getBottomBorder());
			}
			break;
		}
	}
	
	protected void verifyFontColor(Color color, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertEquals(color, c.getFontColor());
		}
	}
	
	protected void verifyFillColor(Color color, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertEquals(color, c.getFillColor());
		}
	}
	
	protected void verifyAlign(AlignType align, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			switch (align) {
			case VERTICAL_ALIGN_TOP:
				Assert.assertEquals("top", c.getVerticalAlign());
				break;
			case VERTICAL_ALIGN_MIDDLE:
				Assert.assertEquals("middle", c.getVerticalAlign());
				break;
			case VERTICAL_ALIGN_BOTTOM:
				Assert.assertEquals("bottom", c.getVerticalAlign());
				break;
			case HORIZONTAL_ALIGN_LEFT:
				Assert.assertEquals("left", c.getHorizontalAlign());
				break;
			case HORIZONTAL_ALIGN_CENTER:
				Assert.assertEquals("center", c.getHorizontalAlign());
				break;
			case HORIZONTAL_ALIGN_RIGHT:
				Assert.assertEquals("right", c.getHorizontalAlign());
				break;
			}
		}
	}
	protected void verifyInsert(Insert type, CellCacheAggeration src, CellCacheAggeration.Builder builder) {
		int tRow = src.getTop();
		int lCol = src.getLeft();
		int bRow = src.getBottom();
		int rCol = src.getRight();
		
		switch (type) {
		case CELL_RIGHT:
			//move cells right
			Assert.assertEquals(src, builder.right().build());
			
			//original cell shall be empty
			verifyCellType(Cell.CellType.BLANK, tRow, lCol, bRow, rCol);
			break;
		case CELL_DOWN:
			//move cells down
			Assert.assertEquals(src, builder.down().build());
			
			//original cell shall be empty
			verifyCellType(Cell.CellType.BLANK, tRow, lCol, bRow, rCol);
			break;
		}
	}
	
	protected void verifyClearContent(CellCacheAggeration cache) {
		
		for (CellCache c : cache) {
			int row = c.getRow();
			int col = c.getCol();
			Cell current = getCell(row, col);
			//style remains
			Assert.assertEquals(c.getBottomBorder(), current.getBottomBorder());
			Assert.assertEquals(c.getRightBorder(), current.getRightBorder());
			Assert.assertEquals(c.getFillColor(), current.getFillColor());
			
			Assert.assertEquals(Cell.CellType.BLANK, current.getCellType());
			Assert.assertEquals("", current.getText());
			Assert.assertEquals("", current.getEdit());
		}
	}
	
	protected void verifyCellType(Cell.CellType cellType, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertEquals(cellType, c.getCellType());
		}
	}
	
	protected void verifyDelete(Delete type, CellCacheAggeration src, CellCacheAggeration.Builder builder, Integer expandSize) {
		switch (type) {
		case CELL_LEFT:
			Assert.assertEquals(src, expandSize == null ? builder.build() : builder.expandDown(expandSize).build());
			break;
		case CELL_UP:
			Assert.assertEquals(src, expandSize == null ? builder.build() : builder.expandRight(expandSize).build());
			break;
		}
	}
	
	protected void verifyClearAll(int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			if (c.getRow() != bRow)
				Assert.assertFalse(c.hasBottomBorder());
			if (c.getCol() != rCol)
				Assert.assertFalse(c.hasRightBorder());
			Assert.assertEquals(new Color("#FFFFFF"), c.getFillColor());
			Assert.assertEquals(Cell.CellType.BLANK, c.getCellType());
			Assert.assertEquals("", c.getText());
			Assert.assertEquals("", c.getEdit());
		}
	}
	
	protected void verifyClearStyle(CellCacheAggeration cache) {
		int bRow = cache.getBottom();
		int rCol = cache.getRight();
		
		for (CellCache c : cache) {
			int row = c.getRow();
			int col = c.getCol();
			Cell current = getCell(row, col);
			if (row != bRow)
				Assert.assertFalse(current.hasBottomBorder());
			if (col != rCol)
				Assert.assertFalse(current.hasRightBorder());
			Assert.assertEquals(new Color("#FFFFFF"), current.getFillColor());
			
			//cellType/edit remain the same, note: text may change, since clear style also means clear format
			Assert.assertEquals(c.getCellType(), current.getCellType());
			Assert.assertEquals(c.getEdit(), current.getEdit());
		}
	}
	
	/**
	 * @param first
	 */
	protected void rightClick(JQuery target) {
		mouseDirector.rightClick(target);
	}
	
	protected void verifyFontFamily(String fontFamily, int tRow, int lCol, int bRow, int rCol) {
		for (int r = tRow; r <= bRow; r++) {
			for (int c = lCol; c <= rCol; c++) {
				Assert.assertEquals(fontFamily, getCell(r, c).getFontFamily());
			}
		}
	}
	
	protected void verifyFontSize(int fontSizeIn, int tRow, int lCol, int bRow, int rCol) {
		for (int r = tRow; r <= bRow; r++) {
			for (int c = lCol; c <= rCol; c++) {
				String fontSizeInPx = getCell(r, c).getFontSize().replace("px", "");
				Assert.assertEquals(fontSizeIn, pxToPoint(Double.parseDouble(fontSizeInPx)));
			}
		}
	}
	
	protected void verifyFontBold(boolean fontBold, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> cs = iterator(tRow, lCol, bRow, rCol);
		while (cs.hasNext()) {
			Cell c = cs.next();
			Assert.assertEquals(fontBold, c.isFontBold());
		}
	}
	
	protected void verifyFontItalic(boolean fontItalic, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertEquals(fontItalic, c.isFontItalic());
		}
	}
	
	protected void verifyPasteAll(PasteSource type, CellCacheAggeration from, int tRow, int lCol, int bRow, int rCol) {
		verifyPasteAll(type, from, getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build());
	}
	
	protected void verifyPasteAll(PasteSource type, CellCacheAggeration from, CellCacheAggeration pasteTo) {
		Assert.assertEquals(from, pasteTo);
		if (type == PasteSource.CUT) {
			verifyClearAll(from.getTop(), from.getLeft(), from.getBottom(), from.getRight());	
		} else { //copy source remain the same
			Assert.assertEquals(from, getCellCacheAggerationBuilder(from.getTop(), from.getLeft(), from.getBottom(), from.getRight()).build());
		}
	}
	
	protected void verifyFormatNumber(FormatNumber type, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			if (c.getCellType() == CellType.NUMBER) {
				switch (type) {
				case ACCOUNTING:
					Assert.assertTrue(c.getText().indexOf("$") >= 0);
					break;
				}
			}
		}
	}
	
	private static int pxToPoint(double px) {
		return Double.valueOf(px * 72 / 96).intValue(); //assume 96dpi
	}
}