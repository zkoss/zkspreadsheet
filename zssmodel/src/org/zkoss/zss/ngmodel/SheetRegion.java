package org.zkoss.zss.ngmodel;

import java.io.Serializable;

import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.poi.ss.util.CellReference;
/**
 * Indicates a region of cells in a sheet
 * @author Dennis
 * @since 3.5.0
 */
public class SheetRegion implements Serializable{
	private static final long serialVersionUID = 1L;

	final NSheet sheet;
	final CellRegion region;
	
	public SheetRegion(NSheet sheet,CellRegion region){
		this.sheet = sheet;
		this.region = region;
	}
	public SheetRegion(NSheet sheet,int row, int column){
		this(sheet,new CellRegion(row,column));
	}
	public SheetRegion(NSheet sheet,int row, int column, int lastRow, int lastColumn){
		this(sheet,new CellRegion(row,column,lastRow,lastColumn));
	}
	public SheetRegion(NSheet sheet,String areaReference){
		
		//regard to testcase 439, now range should support "1:2" -> means row 1 to 2 
		AreaReference ref = new AreaReference(areaReference);
		int row = ref.getFirstCell().getRow();
		int column = ref.getFirstCell().getCol();
		int lastRow = ref.getLastCell().getRow();
		int lastColumn = ref.getLastCell().getCol();
		
		if(row==-1){
			row = 0;
		}
		if(lastRow==-1){
			lastRow = sheet.getBook().getMaxRowIndex();
		}
		if(column==-1){
			column = 0;
		}
		if(lastColumn==-1){
			lastColumn = sheet.getBook().getMaxColumnIndex();
		}
		this.sheet = sheet;
		this.region = new CellRegion(Math.min(row, lastRow), Math.min(column,
				lastColumn), Math.max(row, lastRow), Math.max(column,
				lastColumn));
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
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((region == null) ? 0 : region.hashCode());
		result = prime * result + ((sheet == null) ? 0 : sheet.hashCode());
		return result;
	}
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		SheetRegion other = (SheetRegion) obj;
		if (region == null) {
			if (other.region != null)
				return false;
		} else if (!region.equals(other.region))
			return false;
		if (sheet == null) {
			if (other.sheet != null)
				return false;
		} else if (!sheet.equals(other.sheet))
			return false;
		return true;
	}
	
	public String getReferenceString(){
		if(region.isSingle()){
			return new CellReference(sheet.getSheetName(),region.getRow(), region.getColumn(),false,false).formatAsString();
		}else{
			return new AreaReference(new CellReference(sheet.getSheetName(),region.getRow(), region.getColumn(),false,false), 
				new CellReference(sheet.getSheetName(), region.getLastRow(),region.getLastColumn(),false,false)).formatAsString();
		}
	}
	
}
