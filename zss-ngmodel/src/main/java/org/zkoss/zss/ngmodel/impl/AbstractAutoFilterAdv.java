package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import org.zkoss.zss.ngmodel.NAutoFilter;


public abstract class AbstractAutoFilterAdv implements NAutoFilter,Serializable{
	private static final long serialVersionUID = 1L;

	public static class FilterColumnImpl implements NFilterColumn, Serializable{
		private static final long serialVersionUID = 1L;
		int index;
		List<String> filters;
		Set criteria1;
		Set criteria2;
		boolean on = false;
		FilterOp op = FilterOp.AND;
		
		public FilterColumnImpl(int index){
			this.index = index;
		}
		
		@Override
		public int getIndex() {
			return index;
		}

		@Override
		public List<String> getFilters() {
			return filters==null?Collections.EMPTY_LIST:Collections.unmodifiableList(filters);
		}

		@Override
		public Set getCriteria1() {
			return criteria1==null?Collections.EMPTY_SET:Collections.unmodifiableSet(criteria1);
		}

		@Override
		public Set getCriteria2() {
			return criteria2==null?Collections.EMPTY_SET:Collections.unmodifiableSet(criteria2);
		}

		@Override
		public boolean isOn() {
			return on;
		}

		@Override
		public FilterOp getOperator() {
			return op;
		}

		@Override
		public void addFilter(String filter) {
			if(filters==null){
				filters = new LinkedList<String>();
			}
			filters.add(filter);
		}

		@Override
		public void clearFilter() {
			filters = null;
		}

		@Override
		public void addCriteria1(Object obj) {
			if(criteria1==null){
				criteria1 = new LinkedHashSet();
			}
			criteria1.add(obj);
		}

		@Override
		public void clearCriteria1() {
			criteria1 = null;
		}

		@Override
		public void addCriteria2(Object obj) {
			if(criteria2==null){
				criteria2 = new LinkedHashSet();
			}
			criteria2.add(obj);
		}

		@Override
		public void clearCriteria2() {
			criteria2 = null;
		}

		@Override
		public void setOn(boolean on) {
			this.on = on;
		}

		@Override
		public void setOperator(FilterOp op) {
			this.op = op;
		}
		
	}
}
