/* Cell.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 1, 2012 2:48:41 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.WebDriver;
import org.zkoss.test.Border;
import org.zkoss.test.Color;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.Util;
import org.zkoss.test.Widget;

import com.google.common.base.Objects;
import com.google.inject.Inject;
import com.google.inject.assistedinject.Assisted;

/**
 * @author sam
 *
 */
public class Cell extends Widget {

	public enum CellType {
		NUMBER,
		STRING,
		FORMULA,
		BLANK,
		BOOLEAN,
		ERROR
	}
	
	final int row;
	
	final int col;
	
	final Cell.Factory cellFactory;
	
	@Inject
	/*package*/ Cell(@Assisted("row") Integer row, @Assisted("col") Integer col,
			MainBlock mainBlock, JQueryFactory jqFactory, Cell.Factory cellFactory,
			ConditionalTimeBlocker au, WebDriver webDriver) {
		super(mainBlock.widgetScript() + ".getCell(" + row + "," + col + ")", 
				jqFactory, au, webDriver);
		
		this.row = row;
		this.col = col;
		this.cellFactory = cellFactory;
	}
	
	/**
	 * @return
	 */
	public int getRow() {
		return row;
	}
	
	/**
	 * @return
	 */
	public int getCol() {
		return col;
	}
	
	public CellType getCellType() {
		 int cellType = Util.intValue(javascriptExecutor.executeScript("return " + widgetScript() + ".cellType"));
		 switch (cellType) {
		 case 0:
			 return CellType.NUMBER;
		 case 1:
			 return CellType.STRING;
		 case 2:
			 return CellType.FORMULA;
		 case 3:
			 return CellType.BLANK;
		 case 4:
			 return CellType.BOOLEAN;
		 case 5:
			 return CellType.ERROR;
		 }
		 throw new IllegalArgumentException("Illegal cellType");
	}
	
	public String getText() {
		String text = (String) javascriptExecutor.executeScript("return " +  widgetScript() + ".text");
		return text;
	}
	
	public String getPureText() {
		String text = (String) javascriptExecutor.executeScript("return " +  widgetScript() + ".getPureText()");
		return text;
	}
	
	public String getEdit() {
		String edit = (String) javascriptExecutor.executeScript("return " +  widgetScript() + ".edit");
		return edit;
	}

	/**
	 * @return
	 */
	public boolean isWrap() {
		//TODO: shall also check style
		return (Boolean) javascriptExecutor.executeScript("return " + widgetScript() + ".wrap");
	}
	
	public boolean isMerged() {
		String script = "return " + widgetScript() + ".merid";
		Object hasMergeId = javascriptExecutor.executeScript(script);
		return hasMergeId != null;
	}

	public String getHorizontalAlign() {
		return jq$n("cave").css("text-align");
	}
	
	public String getVerticalAlign() {
		//TODO: IE6/IE7
		return jq$n("cave").css("vertical-align");
	}
	
	public String getFontFamily() {
		String fontFamily = jq$n("real").css("font-family").toLowerCase();
		fontFamily = fontFamily.replace("'", "");//Note. chrome font name may have '
		return fontFamily;
	}
	
	public String getFontSize() {
		return jq$n("real").css("font-size").toLowerCase();
	}
	
	public String getFontWeight() {
		return jq$n("real").css("font-weight").toLowerCase();
	}
	
	public boolean isFontBold() {
		String fontWeight = getFontWeight();
		return "700".equals(fontWeight) || "bold".equalsIgnoreCase(fontWeight);
	}
	
	public boolean isFontItalic() {
		return "italic".equals(jq$n("real").css("font-style").toLowerCase());
	}
	
	public boolean isFontUnderline() {
		return "underline".equals(jq$n("cave").css("text-decoration").toLowerCase());
	}
	
	public boolean isFontStrike() {
		return "line-through".equals(jq$n("cave").css("text-decoration").toLowerCase());
	}
	
	public Color getFontColor() {
		return new Color(jq$n("real").css("color"));
	}
	
	public Color getFillColor() {
		String c = jq$n().css("background-color").toUpperCase();
		//Note. in chrome will use RGBA(0, 0, 0, 0) = TRANSPARENT
		if (c.indexOf("TRANSPARENT") >= 0 || c.indexOf("RGBA(0, 0, 0, 0)") >= 0)
			c = "#FFFFFF";
		return new Color(c);
	}
	
	public Border getBottomBorder() {
		JQuery $n = jq$n();
		return new Border($n.css("border-bottom-width"), $n.css("border-bottom-style"), $n.css("border-bottom-color"));
	}
	
	public boolean hasBottomBorder() {
		Color bottomCellBackgroundColor = cellFactory.create(row + 1, col).getFillColor();
		Border b = getBottomBorder();
		//if cell bottom border color = bottom cell's background color, means this cell doesn't have bottom border
		return !("1px".equals(b.getWidth()) 
			&& "solid".equals(b.getStyle())
			&& Objects.equal(bottomCellBackgroundColor, b.getColor()));
	}
	
	public Border getRightBorder() {
		JQuery $n = jq$n();
		return new Border($n.css("border-right-width"), $n.css("border-right-style"), $n.css("border-right-color"));
	}
	
	public boolean hasRightBorder() {
		Color rightCellBackgroundColor = cellFactory.create(row, col + 1).getFillColor();
		Border b = getRightBorder();
		//if cell right border = right cell's background color, means this cell doesn't have right border
		return !("1px".equals(b.getWidth()) 
			&& "solid".equals(b.getStyle())
			&& Objects.equal(rightCellBackgroundColor, b.getColor()));
	}
	
	@Override
	public String toString() {
		return Objects.toStringHelper(this)
			.add("row", row)
			.add("col", col)
			.toString();
	}



	public static interface Factory {
		public Cell create(@Assisted("row") Integer row, @Assisted("col") Integer col);
	}
}