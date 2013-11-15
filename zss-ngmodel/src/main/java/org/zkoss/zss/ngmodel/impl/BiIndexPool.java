package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.NavigableMap;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

public class BiIndexPool<T> {
	TreeMap<Integer,T> objs = new TreeMap<Integer,T>();
	HashMap<T,Integer> objsReverse = new HashMap<T,Integer>(); 
	
	public T get(int idx) {
		return objs.get(idx);
	}
	
	public T put(int idx, T obj){
		T old = objs.put(idx, obj);
		if(old!=null){
			objsReverse.remove(old);
		}
		objsReverse.put(obj, idx);
		return old;
	}
	
	public int get(T obj){
		Integer idx = objsReverse.get(obj);
		return idx==null?-1:idx;
	}
	
	public int firstKey(){
		return objs.size()<=0?-1:objs.firstKey();
	}
	
	public int lastKey(){
		return objs.size()<=0?-1:objs.lastKey();
	}
	
	public void clear(int start,int end){
		
		NavigableMap<Integer, T> effected = objs.subMap(start,true,end,true);
		
		Iterator<Integer> iter = effected.keySet().iterator();
		while(iter.hasNext()){
			objsReverse.remove(iter.next());
			iter.remove();
		}
	}
	
	public void insert(int start, int size) {
		if(size<=0) return;

		//get last, reversed cell
		SortedMap<Integer,T> effected = objs.descendingMap().headMap(start,true);
		
		//shift from the end
		for(Entry<Integer,T> entry:new ArrayList<Entry<Integer,T>>(effected.entrySet())){
			int idx = entry.getKey();
			int newidx = idx+size;
			T cell = entry.getValue();
			
			objs.remove(idx);
			objsReverse.remove(cell);
			
			objs.put(newidx, cell);
			objsReverse.put(cell, newidx);
		}
	}
	
	public void delete(int start, int size) {
		if(size<=0) return;
		//get last,
		SortedMap<Integer,T> effected = objs.tailMap(start,true);
		
		//shift
		for(Entry<Integer,T> entry:new ArrayList<Entry<Integer,T>>(effected.entrySet())){
			int idx = entry.getKey();
			int newidx = idx-size;
			T cell = entry.getValue();
			objs.remove(idx);
			objsReverse.remove(cell);
			if(newidx>=start){
				objs.put(newidx, cell);
				objsReverse.put(cell, newidx);
			}
		}
	}
	
	public Set<Integer> keySet(){
		return objs.keySet();
	}

	public Collection<T> values() {
		return objs.values();
	}
	
	public Collection<T> subValues(int fromKey,int toKey){
		return objs.subMap(fromKey, true, toKey, true).values();
	}
	
	
}
