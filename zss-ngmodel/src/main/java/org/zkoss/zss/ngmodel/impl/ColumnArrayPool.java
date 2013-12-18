package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collection;
import java.util.SortedMap;
import java.util.TreeMap;

/**
 * to lo/high (firstIdx/lastIdx) index the column array
 * @author dennis
 *
 */
public class ColumnArrayPool implements Serializable{

	private static final long serialVersionUID = 1L;
	private final TreeMap<Integer,ColumnArrayAdv> columnArrayFirst = new TreeMap<Integer,ColumnArrayAdv>();
	private final TreeMap<Integer,ColumnArrayAdv> columnArrayLast = new TreeMap<Integer,ColumnArrayAdv>();
	
	public ColumnArrayPool(){
		
	}

	public boolean hasLastKey(int columnIdx) {
		return columnArrayLast.size()<=0 || columnIdx>columnArrayLast.lastKey();
	}

	public SortedMap<Integer, ColumnArrayAdv> lastSubMap(int columnIdx) {
		return columnArrayLast.subMap(columnIdx, true, columnArrayLast.lastKey(),true);
	}

	public Collection<ColumnArrayAdv> values() {
		return columnArrayLast.values();
	}

	public ColumnArrayAdv overlap(int index, int lastIndex) {
		SortedMap<Integer,ColumnArrayAdv> overlap = columnArrayFirst.size()==0?null:columnArrayFirst.subMap(index, true, lastIndex,true); 
		if(overlap!=null && overlap.size()>0){
			return overlap.get(overlap.firstKey());
		}
		overlap = columnArrayLast.size()==0?null:columnArrayLast.subMap(index, true, lastIndex, true); 
		if(overlap!=null && overlap.size()>0){
			return overlap.get(overlap.firstKey());
		}
		return null;
		
	}

	public int size() {
		return columnArrayLast.size();
	}

	public int lastLastKey() {
		return columnArrayLast.lastKey();
	}

	public void put(ColumnArrayAdv array) {
		ColumnArrayAdv old;
		if((old = columnArrayFirst.put(array.getIndex(),array))!=null){
			throw new IllegalStateException("try to replace a column array in first map "+old +", new "+array);
		}
		if((old = columnArrayLast.put(array.getLastIndex(),array))!=null){
			throw new IllegalStateException("try to replace a column array in last map"+old +", new "+array);
		}
	}

	public void remove(ColumnArrayAdv array) {
		columnArrayFirst.remove(array.getIndex());
		columnArrayLast.remove(array.getLastIndex());
	}

	public int firstFirstKey() {
		return columnArrayFirst.firstKey();
	}

	public void clear() {
		columnArrayFirst.clear();
		columnArrayLast.clear();
	}
}
