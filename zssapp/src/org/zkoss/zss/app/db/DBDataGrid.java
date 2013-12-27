package org.zkoss.zss.app.db;

import java.util.Date;
import java.util.Iterator;

import org.zkoss.zss.app.db.DataSource.RowData;
import org.zkoss.zss.ngmodel.DefaultDataGrid;
import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.NDataCell;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NDataRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;

public class DBDataGrid extends DefaultDataGrid {

	public DBDataGrid(NSheet sheet) {
		super(sheet);
	}

	int cachedIndex = -1;
	RowData cachedRow;
	
	DataSource src = new DataSourceSimpleImpl();
	
	@Override
	public NCellValue getValue(int row, int column) {
		if(column >= 5){
			return getLocalValue(row, column);
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
		if(column >= 5){
			setLocalValue(row, column, value);
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
		
		if(column>=5){
			return true;
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
	public boolean supportOperations() {
		return false;
	}

	@Override
	public void insertRow(int rowIdx, int size) {}

	@Override
	public void deleteRow(int rowIdx, int size) {}

	@Override
	public void insertColumn(int rowIdx, int size) {}

	@Override
	public void deleteColumn(int rowIdx, int size) {}

	@Override
	public boolean supportDataIterator() {
		return false;
	}

	@Override
	public Iterator<NDataRow> getRowIterator() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Iterator<NDataCell> getCellIterator(int row) {
		// TODO Auto-generated method stub
		return null;
	}

}
