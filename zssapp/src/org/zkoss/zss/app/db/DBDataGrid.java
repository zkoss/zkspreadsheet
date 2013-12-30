package org.zkoss.zss.app.db;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.zkoss.zss.app.db.DataSource.RowData;
import org.zkoss.zss.ngmodel.DefaultDataGrid;
import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.NDataCell;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NDataRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;

public class DBDataGrid extends DefaultDataGrid {

	public DBDataGrid() {
	}

	
	
	int cachedIndex = -1;
	RowData cachedRow;
	
	DataSource src = new DataSourceSimpleImpl();
	int maxRow = src.getMaxRow();
	int maxColumn = src.getMaxColumn();
	
	@Override
	public NCellValue getValue(int row, int column) {
		if(column >= maxColumn){
			return super.getValue(row, column);
		}
		if(row>=maxRow){
			return null;
		}
		
		if(cachedIndex!=row){
			cachedIndex = row;
			cachedRow = src.getRowData(row);
		}
		if(cachedRow==null){
			return null;
		}

		
		
		Object[] fields = cachedRow.getFields();
		Object field = fields[column];
		
		if(field instanceof String){
			return new NCellValue((String)field);
		}else if(field instanceof Number){
			return new NCellValue(((Number)field).doubleValue());
		}else if (field instanceof Date){
			return new NCellValue(EngineFactory.getInstance().getCalendarUtil().dateToDoubleValue((Date)field,false));
		}else if(field instanceof Boolean){
			return new NCellValue((Boolean)field);
		}
		return null;
	}

	@Override
	public void setValue(int row, int column, NCellValue value) {
		if(column >= maxColumn){
			super.setValue(row, column, value);
			return;
		}
		if(row>=maxRow){
			return;
		}
		
		if(cachedIndex!=row){
			cachedIndex = row;
			cachedRow = src.getRowData(row);
		}
		if(cachedRow==null){
			return;
		}
		
		Object[] fields = cachedRow.getFields();
		
		
		
		switch(column){
		case 0:
			fields[column] = value.getValue();
			break;
		case 1:
			fields[column] = EngineFactory.getInstance().getCalendarUtil().doubleValueToDate(((Number)value.getValue()).doubleValue());
			break;
		case 2:
		case 3:
			fields[column] = value.getValue();
			break;
		case 4:
			fields[column] = value.getValue();
		}
		
	}

	@Override
	public boolean validateValue(int row, int column, NCellValue value) {
		
		if(column>=maxColumn){
			return true;
		}
		if(row>=maxRow){
			return false;
		}
		if(value==null)
			return false;
		
		Object val = value.getValue();
		switch(column){
		case 0:
			return val instanceof String;
		case 1:
		case 2:
		case 3:
			return val instanceof Number;
		case 4:
			return val instanceof Boolean;
		}
		return true;
	}

	@Override
	public boolean isSupportedOperations() {
		return false;
	}

	@Override
	public boolean isProvidedIterator() {
		return true;
	}

	@Override
	public Iterator<NDataRow> getRowIterator() {
		return new Iterator<NDataRow>(){
			
			int idx = 0;

			@Override
			public boolean hasNext() {
				return idx<maxRow;
			}

			@Override
			public NDataRow next() {
				RowData data = src.getRowData(idx);
				return new DataRowImpl(idx++,data);
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException("readonly");
			}
			
		};
	}

	@Override
	public Iterator<NDataCell> getCellIterator(int row) {
		RowData data = src.getRowData(row);
		return new DataRowImpl(row,data).getCellIterator();
	}
	
	static class DataRowImpl implements NDataRow{

		int index;
		RowData data;
		List<NDataCell> cells;
		public DataRowImpl(int idx, RowData data){
			this.index = idx;
			this.data = data;
		}
		
		@Override
		public Iterator<NDataCell> getCellIterator() {
			if(cells==null){
				cells =  new ArrayList<NDataCell>();
				int i=0;
				for(Object obj:data.getFields()){
					cells.add(new DataCellImpl(index,i++,obj));
				}
			}
			return cells.iterator();
		}

		@Override
		public int getIndex() {
			return index;
		}
		
	}
	
	static class DataCellImpl implements NDataCell{

		int rowIndex;
		int columnIndex;
		Object value;
		
		public DataCellImpl(int rowIdx,int columnIdx,Object value){
			this.rowIndex = rowIdx;
			this.columnIndex = columnIdx;
			this.value = value;
					
		}
		
		@Override
		public NCellValue getValue() {
			if(value instanceof String){
				return new NCellValue((String)value);
			}else if(value instanceof Number){
				return new NCellValue(((Number)value).doubleValue());
			}else if (value instanceof Date){
				return new NCellValue(EngineFactory.getInstance().getCalendarUtil().dateToDoubleValue((Date)value,false));
			}else if(value instanceof Boolean){
				return new NCellValue((Boolean)value);
			}
			return null;
		}

		@Override
		public int getRowIndex() {
			return rowIndex;
		}

		@Override
		public int getColumnIndex() {
			return columnIndex;
		}
		
	}

}
