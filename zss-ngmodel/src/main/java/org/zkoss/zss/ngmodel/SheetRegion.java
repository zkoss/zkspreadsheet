package org.zkoss.zss.ngmodel;

import java.io.Serializable;

public class SheetRegion implements Serializable{
	private static final long serialVersionUID = 1L;

	final NSheet sheet;
	final CellRegion region;
	
	public SheetRegion(NSheet sheet,CellRegion region){
		this.sheet = sheet;
		this.region = region;
	}
	public SheetRegion(NSheet sheet,int row, int column, int lastRow, int lastColumn){
		this(sheet,new CellRegion(row,column,lastRow,lastColumn));
	}
	public SheetRegion(NSheet sheet,String areaReference){
		this(sheet,new CellRegion(areaReference));
	}
	
	public String toString() {
		StringBuilder sb = new StringBuilder();
		sb.append(sheet.getSheetName()).append("!").append(region.toString());
		return sb.toString();
	}
	
	public NSheet getSheet(){
		return sheet;
	}
	
	public CellRegion getRegion(){
		return region;
	}
	
	public int getRow() {
		return region.row;
	}

	public int getColumn() {
		return region.column;
	}

	public int getLastRow() {
		return region.lastRow;
	}

	public int getLastColumn() {
		return region.lastColumn;
	}
	
	public int getRowCount(){
		return region.getRowCount();
	}
	public int getColumnCount(){
		return region.getColumnCount();
	}
	
}
