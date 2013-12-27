package org.zkoss.zss.app.db;

public interface DataSource {

	public static interface RowData{
		public Object getKey();
		public Object[] getFields();
	}
	
	public RowData getRowData(int index);
	
	public boolean canUpdate(Object key, int index, Object field);

	public void update(Object key, int index, Object field);
	
	
}
