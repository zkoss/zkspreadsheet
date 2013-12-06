/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import java.io.Serializable;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NViewAnchor implements Serializable {

	private int rowIndex;
	private int columnIndex;
	private int xOffset;
	private int yOffset;
	private int width;
	private int height;
	public NViewAnchor(int rowIndex, int columnIndex, int width, int height) {
		this(rowIndex,columnIndex,0,0,width,height);
		
	}
	public NViewAnchor(int rowIndex, int columnIndex, int xOffset, int yOffset,
			int width, int height) {
		this.rowIndex = rowIndex;
		this.columnIndex = columnIndex;
		this.xOffset = xOffset;
		this.yOffset = yOffset;
		this.width = width;
		this.height = height;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	public void setColumnIndex(int columnIndex) {
		this.columnIndex = columnIndex;
	}

	public int getXOffset() {
		return xOffset;
	}

	public void setXOffset(int xOffset) {
		this.xOffset = xOffset;
	}

	public int getYOffset() {
		return yOffset;
	}

	public void setYOffset(int yOffset) {
		this.yOffset = yOffset;
	}

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}

}
