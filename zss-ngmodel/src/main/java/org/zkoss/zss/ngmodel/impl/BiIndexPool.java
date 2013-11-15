package org.zkoss.zss.ngmodel.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
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
	
	public Collection<T> clear(int start,int end){
		
		NavigableMap<Integer, T> effected = objs.subMap(start,true,end,true);
		LinkedList<T> remove = new LinkedList<T>(); 
		Iterator<Entry<Integer,T>> iter = effected.entrySet().iterator();
		while(iter.hasNext()){
			Entry<Integer,T> entry = iter.next();
			T obj = entry.getValue();
			objsReverse.remove(obj);
			remove.add(obj);
			iter.remove();
		}
		return remove;
	}
	
	public void insert(int start, int size) {
		if(size<=0) return;

		//get last, reversed cell
		SortedMap<Integer,T> effected = objs.descendingMap().headMap(start,true);
		
		//shift from the end
		for(Entry<Integer,T> entry:new ArrayList<Entry<Integer,T>>(effected.entrySet())){
			int idx = entry.getKey();
			int newidx = idx+size;
			T obj = entry.getValue();
			
			objs.remove(idx);
			objsReverse.remove(obj);
			
			objs.put(newidx, obj);
			objsReverse.put(obj, newidx);
		}
	}
	
	public Collection<T> delete(int start, int size) {
		if(size<=0) return Collections.EMPTY_LIST;
		//get last,
		SortedMap<Integer,T> effected = objs.tailMap(start,true);
		LinkedList<T> remove = new LinkedList<T>(); 
		//shift
		for(Entry<Integer,T> entry:new ArrayList<Entry<Integer,T>>(effected.entrySet())){
			int idx = entry.getKey();
			int newidx = idx-size;
			T obj = entry.getValue();
			objs.remove(idx);
			objsReverse.remove(obj);
			if(newidx>=start){
				objs.put(newidx, obj);
				objsReverse.put(obj, newidx);
			}else{
				remove.add(obj);
			}
		}
		return remove;
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
