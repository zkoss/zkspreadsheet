package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.LinkedList;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.Map.Entry;

/**
 * to lo/high (firstIdx/lastIdx) index the column array
 * @author dennis
 *
 */
/*package*/ class ColumnArrayPool implements Serializable{

	private static final long serialVersionUID = 1L;
	private final TreeMap<Integer,AbstractColumnArrayAdv> columnArrayFirst = new TreeMap<Integer,AbstractColumnArrayAdv>();
	private final TreeMap<Integer,AbstractColumnArrayAdv> columnArrayLast = new TreeMap<Integer,AbstractColumnArrayAdv>();
	
	public ColumnArrayPool(){
		
	}

	public boolean hasLastKey(int columnIdx) {
		return columnArrayLast.size()<=0 || columnIdx>columnArrayLast.lastKey();
	}

	public SortedMap<Integer, AbstractColumnArrayAdv> lastSubMap(int columnIdx) {
		return columnArrayLast.subMap(columnIdx, true, columnArrayLast.lastKey(),true);
	}
	
	public Collection<AbstractColumnArrayAdv> firstSubValues(int start, int end){
		return columnArrayFirst.subMap(start, true, end, true).values();
	}

	public Collection<AbstractColumnArrayAdv> values() {
		return columnArrayLast.values();
	}

	public AbstractColumnArrayAdv overlap(int index, int lastIndex) {
		SortedMap<Integer,AbstractColumnArrayAdv> overlap = columnArrayFirst.size()==0?null:columnArrayFirst.subMap(index, true, lastIndex,true); 
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

	public void put(AbstractColumnArrayAdv array) {
		AbstractColumnArrayAdv old;
		if((old = columnArrayFirst.put(array.getIndex(),array))!=null){
			throw new IllegalStateException("try to replace a column array in first map "+old +", new "+array);
		}
		if((old = columnArrayLast.put(array.getLastIndex(),array))!=null){
			throw new IllegalStateException("try to replace a column array in last map"+old +", new "+array);
		}
	}

	public void remove(AbstractColumnArrayAdv array) {
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
	
	public Collection<AbstractColumnArrayAdv> trim(int start){
		Collection<AbstractColumnArrayAdv> effected = columnArrayFirst.tailMap(start,true).values();
		LinkedList<AbstractColumnArrayAdv> remove = new LinkedList<AbstractColumnArrayAdv>(); 
		for(AbstractColumnArrayAdv array:new ArrayList<AbstractColumnArrayAdv>(effected)){
			remove(array);
			remove.add(array);
		}
		return remove;
	}
}
