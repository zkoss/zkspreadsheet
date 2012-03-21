/* CellCache.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 5:12:22 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import java.util.HashSet;

import org.zkoss.test.Border;
import org.zkoss.test.Color;
import org.zkoss.test.zss.Cell.CellType;

import com.google.common.base.Objects;
import com.google.inject.Inject;
import com.google.inject.assistedinject.Assisted;

/**
 * @author sam
 *
 */
public class CellCache {

	public enum Field {
		MERGE, 
		TEXT, 
		EDIT,
		HORIZONTAL_ALIGN,
		VERTICAL_ALIGN,
		FONT_COLOR,
		FILL_COLOR,
		BOTTOM_BORDER,
		RIGHT_BORDER
	}
	
	public enum EqualCondition {
//		FORMULA,
		VALUE,
		EXPECT_BORDER,
		COLUMN_WIDTH_ONLY,
		IGNORE_NUMBER_FORMAT
	}
	
	int row;
	
	int col;
	
	int width;
	
	int height;
	
	Cell.CellType cellType;
	
	boolean merge;
	
	String text;
	
	String edit;
	
	String horizontalAlign;
	
	String verticalAlign;
	
	Color fontColor;
	
	Color fillColor;
	
	Border bottomBorder;
	
	Border rightBorder;
	
	final Cell.Factory cellFactory;
	
	
	@Inject
	/*package*/ CellCache(@Assisted("row") Integer row, @Assisted("col") Integer col,
			Cell.Factory cellFactory) {
		this.cellFactory = cellFactory;
		this.row = row;
		this.col = col;
		
		Cell cell = cellFactory.create(row, col);
		width = cell.getWidth();
		height = cell.getHeight();
		cellType = cell.getCellType();
		merge = cell.isMerged();
		text = cell.getText();
		edit = cell.getEdit();
		horizontalAlign = cell.getHorizontalAlign();
		verticalAlign = cell.getVerticalAlign();
		fontColor = cell.getFontColor();
		fillColor = cell.getFillColor();
		bottomBorder = cell.getBottomBorder();
		rightBorder = cell.getRightBorder();
	}
	
	public int getRow() {
		return row;
	}
	
	public int getCol() {
		return col;
	}
	
	public CellType getCellType() {
		return cellType;
	}
	
	public int getWidth() {
		return width;
	}
	
	public int getHeight() {
		return height;
	}
	
	public void setMerge(boolean merge) {
		this.merge = merge;
	}
	
	public boolean getMerge() {
		return this.merge;
	}
	
	public void setText(String text) {
		this.text = text;
	}
	
	public String getText() {
		return text;
	}
	
	public void setEdit(String edit) {
		this.edit = edit;
	}
	
	public String getEdit() {
		return edit;
	}
	
	public void setHorizontalAlign(String horizontalAlign) {
		this.horizontalAlign = horizontalAlign;
	}
	
	public String getHorizontalAlign() {
		return horizontalAlign;
	}
	
	public void setVerticalAlign(String verticalAlign) {
		this.verticalAlign = verticalAlign;
	}
	
	public String getVerticalAlign() {
		return verticalAlign;
	}
	
	public void setFontColor(Color fontColor) {
		this.fontColor = fontColor;
	}
	
	public Color getFontColor() {
		return fontColor;
	}
	
	public void setFillColor(Color fillColor) {
		this.fillColor = fillColor;
	}
	
	public Color getFillColor() {
		return fillColor;
	}
	
	public void setBottomBorder(Border bottomBorder) {
		this.bottomBorder = bottomBorder;
	}
	
	public Border getBottomBorder() {
		return bottomBorder;
	}
	
	public void setRightBorder(Border rightBorder) {
		this.rightBorder = rightBorder;
	}
	
	public Border getRightBorder() {
		return rightBorder;
	}
	
	public boolean equals(CellCache that, EqualCondition... fields) {
		HashSet<EqualCondition> set = tranform(fields);
		
		boolean equal = false;
		if (set.contains(EqualCondition.COLUMN_WIDTH_ONLY)) {
			return this.width == that.width;
		}
		
		if (set.contains(EqualCondition.VALUE)) {
			equal = this.text.equals(that.getText());
			if (this.cellType == CellType.FORMULA) {
				equal = that.cellType != CellType.FORMULA;
			} else if (this.cellType == CellType.NUMBER) {
				equal = Objects.equal(this.edit, that.edit);
				if (!set.contains(EqualCondition.IGNORE_NUMBER_FORMAT)) {
					equal = equal && Objects.equal(this.text, that.text); 
				}
			} else {//string, boolean, blank shall be the same
				equal = Objects.equal(this.text, that.text) && 
					Objects.equal(this.edit, that.edit);
			}
		} else {
			equal = Objects.equal(this.cellType, that.cellType);
			if (this.cellType == CellType.NUMBER) {
				equal = Objects.equal(this.edit, that.edit);
			} else if (this.cellType != CellType.FORMULA) {
				equal = Objects.equal(this.edit, that.edit);
				if (!set.contains(EqualCondition.IGNORE_NUMBER_FORMAT)) {
					equal = equal && Objects.equal(this.text, that.text); 
				}
			}
		}
		if (!equal) {
			return equal;
		}
		
		if (set.contains(EqualCondition.EXPECT_BORDER)) {
			//TODO: fine tune comparison
			//1. get right cell background to compare right border
			//2. get bottom cell's background to compare bottom border
			equal = Objects.equal(this.fillColor, that.rightBorder.getColor());
		} else {
			equal = Objects.equal(this.bottomBorder, that.bottomBorder)
				&& Objects.equal(this.rightBorder, that.rightBorder);
		}
		
		equal = Objects.equal(this.horizontalAlign, that.horizontalAlign)
			&& Objects.equal(this.verticalAlign, that.verticalAlign)
			&& Objects.equal(this.fontColor, that.fontColor)
			&& Objects.equal(this.fillColor, that.fillColor);
		return equal;
	}
	
	private HashSet<EqualCondition> tranform(EqualCondition... from) {
		HashSet<EqualCondition> set = new HashSet<EqualCondition>();
		for (EqualCondition i : from) {
			set.add(i);
		}
		return set;
	}
	
	@Override
	public boolean equals(Object obj) {
		if (obj instanceof CellCache) {
			CellCache that = (CellCache)obj;
			
			boolean equal = true;
			if (this.cellType == CellType.FORMULA) {
				equal = that.cellType == CellType.FORMULA;
			} else {//number/string/blank
				equal = Objects.equal(this.text, that.text)
					&& Objects.equal(this.edit, that.edit);
			}
			equal = equal && Objects.equal(this.cellType, that.cellType) 
//				&& Objects.equal(this.text, that.text)
//				&& (this.cellType == CellType.FORMULA && that.cellType == CellType.FORMULA) || Objects.equal(this.edit, that.edit)
				&& Objects.equal(this.horizontalAlign, that.horizontalAlign)
				&& Objects.equal(this.verticalAlign, that.verticalAlign)
				&& Objects.equal(this.fontColor, that.fontColor)
				&& Objects.equal(this.fillColor, that.fillColor)
				&& Objects.equal(this.bottomBorder, that.bottomBorder)
				&& Objects.equal(this.rightBorder, that.rightBorder);
			if (!equal) {
				//TODO: use logger ?
				System.out.println("this [r:" + this.row + ",c:" + this.col + "] != that [r:" + that.row + ",c:" + that.col + "]");
				System.out.println("this: " + this.toString());
				System.out.println("that: " + that.toString());
				
				System.out.println("text? " + Objects.equal(this.text, that.text));
				System.out.println("edit? " + Objects.equal(this.edit, that.edit));
				System.out.println("horizontalAlign? " + Objects.equal(this.horizontalAlign, that.horizontalAlign));
				System.out.println("verticalAlign? " + Objects.equal(this.verticalAlign, that.verticalAlign));
				System.out.println("fontColor? " + Objects.equal(this.fontColor, that.fontColor));
				System.out.println("fillColor? " + Objects.equal(this.fillColor, that.fillColor));
				System.out.println("bottomBorder? " + Objects.equal(this.bottomBorder, that.bottomBorder));
				System.out.println("rightBorder? " + Objects.equal(this.rightBorder, that.rightBorder));
			}
			return equal;
		}
		return false;
	}

	@Override
	public String toString() {
		return Objects.toStringHelper(this)
			.add("row", row)
			.add("col", col)
			.add("text", text)
			.add("edit", edit)
			.add("merge", merge)
			.add("horizontalAlign", horizontalAlign)
			.add("verticalAlign", verticalAlign)
			.add("fontColor", fontColor)
			.add("fillColor", fillColor)
			.add("rightBorder", rightBorder)
			.add("bottomBorder", bottomBorder)
			.toString();
	}
	
	public static interface Factory {
		public CellCache create(@Assisted("row") Integer row, @Assisted("col") Integer col);
	}
}