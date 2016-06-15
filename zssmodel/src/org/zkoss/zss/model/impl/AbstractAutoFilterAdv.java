/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.Map;
import java.util.SortedSet;

import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SColorFilter;
import org.zkoss.zss.model.SCustomFilters;
import org.zkoss.zss.model.SDynamicFilter;
import org.zkoss.zss.model.STop10Filter;
import org.zkoss.zss.range.impl.FilterRowInfo;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
public abstract class AbstractAutoFilterAdv implements SAutoFilter,Serializable{
	private static final long serialVersionUID = 1L;
	
	public abstract void renameSheet(SBook book, String oldName, String newName); //ZSS-555
	
	/**
	 * @since 3.5.0
	 */
	public static class FilterColumnImpl implements NFilterColumn, Serializable{
		private static final long serialVersionUID = 1L;
		private int _index;
		private List<String> _filters;
		private Set _criteria1;
		private Set _criteria2;
		private Boolean _showButton;
		private FilterOp _op = FilterOp.AND;
		private SColorFilter _colorFilter; //ZSS-1191
		private SCustomFilters _customFilters; //ZSS-1224
		private int type; //ZSS-1192: //Date 0, Number 1, String 2
		private SDynamicFilter _dynamicFilter; //ZSS-1226
		private STop10Filter _top10Filter; //ZSS-1127
		
		//ZSS-1193
		private List<FilterRowInfo> orderedRowInfos;
		private int filterType;
		
		public FilterColumnImpl(int index){
			this._index = index;
		}
		
		@Override
		public int getIndex() {
			return _index;
		}

		@Override
		public List<String> getFilters() {
			return _filters==null?Collections.EMPTY_LIST:Collections.unmodifiableList(_filters);
		}

		@Override
		public Set getCriteria1() {
			return _criteria1==null?Collections.EMPTY_SET:Collections.unmodifiableSet(_criteria1);
		}

		@Override
		public Set getCriteria2() {
			return _criteria2==null?Collections.EMPTY_SET:Collections.unmodifiableSet(_criteria2);
		}

		@Override
		public boolean isShowButton() {
			return _showButton==null?true:_showButton.booleanValue();
		}

		@Override
		public FilterOp getOperator() {
			return _op;
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
			_setProperties(filterOp, criteria1, criteria2, showButton, Collections.EMPTY_MAP);
		}
		
		//ZSS-1191: for import
		@Override
		public void setProperties(FilterOp filterOp, Object criteria1,
				Object criteria2, Boolean showButton, Map<String, Object> extra) {
			_setProperties(filterOp, criteria1, criteria2, showButton, extra);
		}
		
		//ZSS-1191
		private void _setProperties(FilterOp filterOp, Object criteria1,
				Object criteria2, Boolean showButton, Map<String, Object> extra) {
			
			//ZSS-1191
			_colorFilter = (SColorFilter) extra.get("colorFilter");
			
			//ZSS-1224
			_customFilters = (SCustomFilters) extra.get("customFilters");
			
			//ZSS-1226
			SDynamicFilter dynaFilter = (SDynamicFilter) extra.get("dynamicFilter");
			if (!DynamicFilterImpl.NOOP_DYNAFILTER.equals(dynaFilter)) { //ZSS-1193
				_dynamicFilter = dynaFilter;
			}
			
			//ZSS-1227
			_top10Filter = (STop10Filter) extra.get("top10Filter");
			
			this._op = filterOp;
			this._criteria1 = getCriteriaSet(criteria1);
			this._criteria2 = getCriteriaSet(criteria2);
			boolean blank1 = this._criteria1.contains("=");
			if(showButton!=null){
				_showButton = showButton; // ZSS-705
			}
			
			
			
			if (criteria1 == null) { //remove filtering
				_filters = null;
				return;
			}
			
			//TODO, more filtering operation
			switch(filterOp) {
			case VALUES:
				
				_filters =  new LinkedList<String>();
				
				for(Object obj:this._criteria1){
					if(obj instanceof String){
						_filters.add((String)obj);
					}
				}
				if(_filters.size()==0){
					_filters = null;
				}
			}
		}

		//ZSS-688
		//@since 3.6.0
		/*package*/ FilterColumnImpl cloneFilterColumnImpl() { //ZSS-1191
			return cloneFilterColumnImpl(null);
		}
		
		//ZSS-1183, ZSS-1191
		//@since 3.9.0
		/*package*/ FilterColumnImpl cloneFilterColumnImpl(SBook book) { //ZSS-1191
			FilterColumnImpl tgt = new FilterColumnImpl(this._index);
			if (this._filters != null) {
				tgt._filters = new LinkedList<String>();
				for (String f : this._filters) {
					tgt._filters.add(f);
				}
			}
			if (this._criteria1 != null) {
				tgt._criteria1 = new HashSet();
				for (Object c : this._criteria1) {
					tgt._criteria1.add(c);
				}
			}
			if (this._criteria2 != null) {
				tgt._criteria2 = new HashSet();
				for (Object c : this._criteria2) {
					tgt._criteria2.add(c);
				}
			}
			tgt._showButton = this._showButton;
			tgt._op = this._op;
			
			//ZSS-1183, ZSS-1191
			if (this._colorFilter != null) {
				tgt._colorFilter = ((ColorFilterImpl)this._colorFilter).cloneColorFilter(book);
			}
			
			//ZSS-1183, ZSS-1224
			if (this._customFilters != null) {
				tgt._customFilters = ((CustomFiltersImpl)this._customFilters).cloneCustomFilters(); 
			}
			
			//ZSS-1183, ZSS-1226
			if (this._dynamicFilter != null) {
				tgt._dynamicFilter = ((DynamicFilterImpl)this._dynamicFilter).cloneDynamicFilter();
			}
			
			//ZSS-1183, ZSS-1227
			if (this._top10Filter != null) {
				tgt._top10Filter = ((Top10FilterImpl)this._top10Filter).cloneTop10Filter();
			}
			return tgt;
		}
		
		//ZSS-1191
		//@since 3.9.0
		@Override
		public SColorFilter getColorFilter() {
			return _colorFilter;
		}
		
		//ZSS-1224
		//@since 3.9.0
		@Override
		public SCustomFilters getCustomFilters() {
			return _customFilters;
		}
		
		//ZSS-1226
		//@since 3.9.0
		@Override
		public SDynamicFilter getDynamicFilter() {
			return _dynamicFilter;
		}
		
		//ZSS-1227
		//@since 3.9.0
		@Override
		public STop10Filter getTop10Filter() {
			return _top10Filter;
		}
		
		//ZSS-1193
		//@since 3.9.0
		//@Internal
		//@See AutoFilterDefaultHandler
		public void setCachedSet(SortedSet<FilterRowInfo> orderedRowInfos) {
			this.orderedRowInfos = orderedRowInfos == null ? null: 
					new ArrayList<FilterRowInfo>(orderedRowInfos);
		}
		
		//ZSS-1193
		//@since 3.9.0
		//@Internal
		//@See CustomFiltersCtrl
		public List<FilterRowInfo> getCachedSet() {
			return this.orderedRowInfos;
		}
		
		//ZSS-1193
		//@since 3.9.0
		//@Internal
		//@See AutoFitlerDefaultHandler
		//1: Date; 2: Number; 3: Text
		public void setFilterType(int type) {
			this.filterType = type;
		}
		
		//ZSS-1193
		//@since 3.9.0
		//@Internal
		//@See CustomFiltersCtrl
		//1: Date; 2: Number; 3: Text
		public int getFilterType() {
			return this.filterType;
		}
		
		//ZSS-1229
		//@since 3.9.0
		//Whether this FilterColumn filtered out some rows?
		public boolean isFiltered() {
			return (_filters != null && !_filters.isEmpty())
				|| (_colorFilter != null)
				|| (_customFilters != null)
				|| (_dynamicFilter != null)
				|| (_top10Filter != null);
		}
		
		//ZSS-1230
		//@since 3.9.0
		//@internal
		public void setIndex(int index0) {
			this._index = index0;
		}
	}
}
