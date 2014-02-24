package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collections;
import java.util.HashSet;
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
		Boolean showButton;
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
		public boolean isShowButton() {
			return showButton==null?true:showButton.booleanValue();
		}

		@Override
		public FilterOp getOperator() {
			return op;
		}

		private Set getCriteriaSet(Object criteria) {
			Set set = new HashSet();
			if (criteria instanceof String[]) {
				String[] strings = (String[]) criteria;
				for(int j = 0; j < strings.length; ++j) {
					set.add(strings[j]);
				}
			}else if (criteria instanceof Set){
				set = (Set)criteria;
			}
			return set;
		}
		
		@Override
		public void setProperties(FilterOp filterOp, Object criteria1,
				Object criteria2, Boolean showButton) {
			this.op = filterOp;
			this.criteria1 = getCriteriaSet(criteria1);
			this.criteria2 = getCriteriaSet(criteria2);
			boolean blank1 = this.criteria1.contains("=");
			if(showButton!=null){
				showButton = showButton;
			}
			
			
			
			if (criteria1 == null) { //remove filtering
				filters = null;
				return;
			}
			
			//TODO, more filtering operation
			switch(filterOp) {
			case VALUES:
				
				filters =  new LinkedList<String>();
				
				for(Object obj:this.criteria1){
					if(obj instanceof String){
						filters.add((String)obj);
					}
				}
				if(filters.size()==0){
					filters = null;
				}
				
//				final String[] filters = (String[]) criteria1;
				//remove old
//				if (_ctfc.isSetFilters()) {
//					_ctfc.unsetFilters();
//				}
				//TODO zss 3.5 WHAT is this?
//				final CTFilters cflts = _ctfc.addNewFilters();
//				if (blank1) {
//					cflts.setBlank(blank1);
//				}
//				for(int j = 0; j < filters.length; ++j) {
//					final CTFilter cflt = cflts.addNewFilter();
//					cflt.setVal(filters[j]);
//				}
			}
		}
		
		
	}
}
