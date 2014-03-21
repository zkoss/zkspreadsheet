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
package org.zkoss.zss.range.impl;

import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColumn;
import org.zkoss.zss.model.SColumnArray;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.impl.AbstractCellAdv;
import org.zkoss.zss.range.SRange;

/**
 * Helper for set cell, row, column style
 * @author Dennis
 *
 */
public class ClearCellHelper extends RangeHelperBase{

	SCellStyle _defaultStyle;
	
	public ClearCellHelper(SRange range) {
		super(range);
		_defaultStyle = range.getSheet().getBook().getDefaultCellStyle();
	}

	public void clearCellStyle() {
		if(isWholeSheet()){
			clearWholeSheetStyle();
		}else if(isWholeRow()) { 
			clearWholeRowCellStyle();
		}else if(isWholeColumn()){
			clearWholeColumnCellStyle();
		}else{
			for(int r = getRow(); r <= getLastRow(); r++){
				SRow row = sheet.getRow(r);
				for (int c = getColumn(); c <= getLastColumn(); c++){
					SColumn column = sheet.getColumn(c);
					SCell cell = sheet.getCell(r,c);
					if(row.getCellStyle(true)!=null || column.getCellStyle(true)!=null){
						//set cell style to default one if there are row or column style here
						cell.setCellStyle(_defaultStyle);
					}else if(cell.getCellStyle(true)!=null){
						cell.setCellStyle(null);
					}
				}
			}
		}
	}

	private void clearWholeSheetStyle() {
		Iterator<SColumnArray> columns = sheet.getColumnArrayIterator();
		while(columns.hasNext()){
			SColumnArray column = columns.next();
			if(column.getCellStyle(true)!=null){
				column.setCellStyle(null);
			}
		}
		Iterator<SRow> rows = sheet.getRowIterator();
		while(rows.hasNext()){
			SRow row = rows.next();
			if(row.getCellStyle(true)!=null){
				row.setCellStyle(null);
			}
			Iterator<SCell> cells = sheet.getCellIterator(row.getIndex());
			while(cells.hasNext()){
				SCell cell = cells.next();
				if(cell.getCellStyle(true)!=null){
					cell.setCellStyle(null);
				}
			}
		}
	}

	private void clearWholeColumnCellStyle() {
		for(int c = getColumn(); c <= getLastColumn(); c++){
			SColumn column = sheet.getColumn(c);
			if(!column.isNull()){
				column.setCellStyle(null);
			}
		}
		Iterator<SRow> rows = sheet.getRowIterator();
		while(rows.hasNext()){
			SRow row = rows.next();
			boolean rowHasStyle = row.getCellStyle(true)!=null;
			for(int c = getColumn(); c <= getLastColumn(); c++){
				SCell cell = sheet.getCell(row.getIndex(), c);
				if(rowHasStyle){
					cell.setCellStyle(_defaultStyle);
				}else if(cell.getCellStyle(true)!=null){
					cell.setCellStyle(null);
				}
			}
		}
	}

	private void clearWholeRowCellStyle() {
		Set<Integer> columnHasStyle = new HashSet<Integer>();
		Iterator<SColumn> columns = sheet.getColumnIterator();
		while(columns.hasNext()){
			SColumn column = columns.next();
			if(column.getCellStyle(true)!=null){
				columnHasStyle.add(column.getIndex());
			}
		}
		
		for(int r = getRow(); r <= getLastRow(); r++){
			Set<Integer> rowColumnHasStyle = new HashSet(columnHasStyle);
			SRow row = sheet.getRow(r);
			if(!row.isNull()){
				row.setCellStyle(null);
			}
			
			Iterator<SCell> cells = sheet.getCellIterator(r);
			while(cells.hasNext()){
				SCell cell = cells.next();

				if(columnHasStyle.contains(cell.getColumnIndex())){
					cell.setCellStyle(_defaultStyle);
					rowColumnHasStyle.remove(cell.getColumnIndex());
				}else if(cell.getCellStyle(true)!=null){
					cell.setCellStyle(null);
				}
			}
			//there are column has style but no cell is this row, have to create a cell and set the default style to it
			for(Integer c:rowColumnHasStyle){
				sheet.getCell(row.getIndex(), c).setCellStyle(_defaultStyle);
			}
		}
	}
	
	
	public void clearCellContent() {
		if(isWholeSheet()){
			clearWholeSheetContent();
		}else if(isWholeRow()) { 
			clearWholeRowContent();
		}else if(isWholeColumn()){
			clearWholeColumnContent();
		}else{
			for(int r = getRow(); r <= getLastRow(); r++){
				for (int c = getColumn(); c <= getLastColumn(); c++){
					SCell cell = sheet.getCell(r,c);
					clearCellContent(cell);
				}
			}
		}
	}

	private void clearWholeSheetContent() {
		Iterator<SRow> rows = sheet.getRowIterator();
		while(rows.hasNext()){
			SRow row = rows.next();
			Iterator<SCell> cells = sheet.getCellIterator(row.getIndex());
			while(cells.hasNext()){
				SCell cell = cells.next();
				clearCellContent(cell);
			}
		}
	}

	private void clearWholeColumnContent() {
		Iterator<SRow> rows = sheet.getRowIterator();
		while(rows.hasNext()){
			SRow row = rows.next();
			for(int c = getColumn(); c <= getLastColumn(); c++){
				SCell cell = sheet.getCell(row.getIndex(), c);
				clearCellContent(cell);
			}
		}
	}

	private void clearWholeRowContent() {
		for(int r = getRow(); r <= getLastRow(); r++){
			Iterator<SCell> cells = sheet.getCellIterator(r);
			while(cells.hasNext()){
				SCell cell = cells.next();
				clearCellContent(cell);
			}
		}
	}
	
	private void clearCellContent(SCell cell){
		if (!cell.isNull()) {
			cell.setHyperlink(null);
			cell.clearValue();
		}
	}

}
