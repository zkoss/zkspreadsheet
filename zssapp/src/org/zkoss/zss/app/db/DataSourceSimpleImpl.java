package org.zkoss.zss.app.db;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Random;
import java.util.TreeMap;


public class DataSourceSimpleImpl implements DataSource{

	public DataSourceSimpleImpl(){}

	TreeMap<Object,RowDataImpl> data = new TreeMap<Object,RowDataImpl>();
	
	Random random = new Random(System.currentTimeMillis());
	SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd");
	public RowData getRowData(int index){
		RowDataImpl row = data.get(index); 
		if(row==null){
			Object key = index;
			data.put(key, row = new RowDataImpl(key,createFakeFields(index)));
		}
		System.out.println("load row "+row.getKey());
		return row;
	}
	
	private Object[] createFakeFields(int row) {
		Object[] fields = new Object[5];
		
		fields[0] = "User "+row;
		fields[1] = randomDate(); 
		fields[2] = 18+random.nextInt(50);
		fields[3] = random.nextInt(99999999);
		fields[4] = random.nextBoolean();
		return fields;
	}

	private Object randomDate() {
		
		StringBuilder sb = new StringBuilder();
		sb.append(2010+random.nextInt(13)).append("/").append(1+random.nextInt(12)).append("/").append(1+random.nextInt(28));
		try {
			return format.parse(sb.toString());
		} catch (ParseException e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}

	public boolean canUpdate(Object key, int fieldIdx, Object field){
		RowDataImpl row = data.get(key); 
		if(row==null){
			return false;
		}
		switch(fieldIdx){
		case 0:
			return field instanceof String;
		case 1:
			return field instanceof Date;
		case 2:
		case 3:
			return field instanceof Number;
		case 4:
			return field instanceof Boolean;
		}
		return false;
	}

	public void update(Object key, int fieldIdx, Object field){
		RowDataImpl row = data.get(key); 
		if(row==null){
			return;
		}
		Object[] fields = row.getFields();
		fields[fieldIdx] = field;
	}
	
	public static class RowDataImpl implements RowData {
		Object key;
		Object[] fields;

		public RowDataImpl(Object key, Object[] fields) {
			super();
			this.key = key;
			this.fields = fields;
		}

		public Object getKey() {
			return key;
		}

		public Object[] getFields() {
			return fields;
		}

	}
	
}
