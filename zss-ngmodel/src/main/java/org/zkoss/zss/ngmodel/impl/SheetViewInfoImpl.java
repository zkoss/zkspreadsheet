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
package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Iterator;
import java.util.TreeSet;

import org.zkoss.zss.ngmodel.NFooter;
import org.zkoss.zss.ngmodel.NHeader;
import org.zkoss.zss.ngmodel.NSheetViewInfo;

public class SheetViewInfoImpl implements NSheetViewInfo, Serializable {
	private static final long serialVersionUID = 1L;
	private boolean displayGridline = true;
	
	private int rowFreeze = 0;
	private int columnFreeze = 0;

	private NHeader header;
	
	private NFooter footer;
	
	private TreeSet<Integer> rowBreaks;
	private TreeSet<Integer> columnBreaks;
	
	
	@Override
	public boolean isDisplayGridline() {
		return displayGridline;
	}

	@Override
	public void setDisplayGridline(boolean enable) {
		displayGridline = enable;
	}

	@Override
	public int getNumOfRowFreeze() {
		return rowFreeze;
	}

	@Override
	public int getNumOfColumnFreeze() {
		return columnFreeze;
	}

	@Override
	public void setNumOfRowFreeze(int num) {
		rowFreeze = num;
	}

	@Override
	public void setNumOfColumnFreeze(int num) {
		columnFreeze = num;
	}

	@Override
	public NHeader getHeader() {
		if(header==null){
			header = new HeaderFooterImpl();
		}
		return header;
	}

	@Override
	public NFooter getFooter() {
		if(footer==null){
			footer = new HeaderFooterImpl();
		}
		return footer;
	}

	@Override
	public int[] getRowBreaks() {
		if(rowBreaks==null){
			return new int[0];
		}
		int[] arr = new int[]{rowBreaks.size()};
		Iterator<Integer> iter = rowBreaks.iterator();
		for(int i=0;i<arr.length;i++){
			arr[i] = iter.next();
		}
		return arr;
	}

	@Override
	public void setRowBreaks(int[] breaks) {
		if(rowBreaks!=null)
			rowBreaks.clear();
		else
			rowBreaks = new TreeSet<Integer>();
		if(breaks!=null){
			for(int i:breaks){
				rowBreaks.add(i);
			}
		}
	}
	
	public void addRowBreak(int row){
		if(rowBreaks==null)
			rowBreaks = new TreeSet<Integer>();
		
		rowBreaks.add(row);
	}
	
	public void addColumnBreak(int column){
		if(columnBreaks==null)
			columnBreaks = new TreeSet<Integer>();
		
		columnBreaks.add(column);
	}

	@Override
	public int[] getColumnBreaks() {
		if(columnBreaks==null){
			return new int[0];
		}
		int[] arr = new int[]{columnBreaks.size()};
		Iterator<Integer> iter = columnBreaks.iterator();
		for(int i=0;i<arr.length;i++){
			arr[i] = iter.next();
		}
		return arr;
	}

	@Override
	public void setColumnBreaks(int[] breaks) {
		if(rowBreaks!=null)
			columnBreaks.clear();
		else
			columnBreaks = new TreeSet<Integer>();
		if(breaks!=null){
			for(int i:breaks){
				columnBreaks.add(i);
			}
		}
	}
	
	
}
