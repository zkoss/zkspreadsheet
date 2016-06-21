package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.SortedSet;
import java.util.TreeMap;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SAutoFilter;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.util.Validations;
import org.zkoss.zss.range.impl.FilterRowInfo;
/**
 * The auto fitler implement
 * @author dennis
 * @since 3.5.0
 */
public class AutoFilterImpl extends AbstractAutoFilterAdv {
	private static final long serialVersionUID = 1L;
	
	private final CellRegion _region;
	private final TreeMap<Integer,NFilterColumn> _columns;
	private Map<Integer, List<FilterRowInfo>> orderedRowInfosMap;
	private Map<Integer, Integer> filterTypes; //ZSS-1234. 1: Date, 2: Number, 3: String

	public AutoFilterImpl(CellRegion region){
		this._region = region;
		_columns = new TreeMap<Integer,NFilterColumn>();
		orderedRowInfosMap = new HashMap<Integer, List<FilterRowInfo>>();
		filterTypes = new HashMap<Integer, Integer>(); //ZSS-1234
	}
	
	@Override
	public CellRegion getRegion() {
		return _region;
	}

	@Override
	public Collection<NFilterColumn> getFilterColumns() {
		return Collections.unmodifiableCollection(_columns.values());
	}

	@Override
	public NFilterColumn getFilterColumn(int index, boolean create) {
		NFilterColumn col = _columns.get(index);
		if(col==null && create){
			int s = _region.getLastColumn()-_region.getColumn()+1; 
			if(index>=s){
				throw new IllegalStateException("the column index "+index+" >= "+s);
			}
			_columns.put(index, col=new FilterColumnImpl(index));
		}
		return col;
	}

	@Override
	public void clearFilterColumn(int index) {
		_columns.remove(index);
	}
	
	@Override
	public void clearFilterColumns() {
		_columns.clear();
	}	

	//ZSS-555
	@Override
	public void renameSheet(SBook book, String oldName, String newName) {
		Validations.argNotNull(oldName);
		Validations.argNotNull(newName);
		if (oldName.equals(newName)) return; // nothing change, let go
		
		final String bookName = book.getBookName();
		// remove old ObjectRef
		Ref dependent = new ObjectRefImpl(bookName, oldName, "AUTO_FILTER", ObjectType.AUTO_FILTER);

		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		dt.clearDependents(dependent);
		
		// Add new ObjectRef into DependencyTable so we can extend/shrink/move
		dependent = new ObjectRefImpl(bookName, newName, "AUTO_FILTER", ObjectType.AUTO_FILTER);
		
		// prepare new dummy CellRef to enforce DataValidation reference dependency
		if (this._region != null) {
			Ref dummy = new RefImpl(bookName, newName, 
				_region.row, _region.column, _region.lastRow, _region.lastColumn);
			dt.add(dependent, dummy);
		}
	}
	
	//ZSS-688
	//@since 3.6.0
	/*package*/ AutoFilterImpl cloneAutoFilterImpl() {
		return cloneAutofilterImpl(null);
	}
	
	//ZSS-1183, ZSS-1191
	/*package*/ AutoFilterImpl cloneAutofilterImpl(SBook book) {
		final AutoFilterImpl tgt = 
				new AutoFilterImpl(new CellRegion(this._region.row, this._region.column, this._region.lastRow, this._region.lastColumn));

		for (SAutoFilter.NFilterColumn value : this._columns.values()) {
			final FilterColumnImpl srccol = (FilterColumnImpl) value;
			final FilterColumnImpl tgtcol = srccol.cloneFilterColumnImpl(book); 
			tgt._columns.put(tgtcol.getIndex(), tgtcol); //ZSS-1183
		}
		
		return tgt;
	}
	
	//ZSS-1229
	@Override
	public boolean isFiltered() {
		for (SAutoFilter.NFilterColumn value : this._columns.values()) {
			if (value.isFiltered()) {
				return true;
			}
		}
		return false;
	}
	
	//ZSS-1230
	//@since 3.9
	//@Internal
	public void putFilterColumn(int index, NFilterColumn filterColumn) {
		_columns.put(index, filterColumn);
	}
	
	//ZSS-1193, ZSS-1233: should cached in autofilter by column index
	//@since 3.9.0
	//@Internal
	//@See AutoFilterDefaultHandler
	public void setCachedSet(int index, SortedSet<FilterRowInfo> orderedRowInfos) {
		this.orderedRowInfosMap.put(index, orderedRowInfos == null ? null: 
				new ArrayList<FilterRowInfo>(orderedRowInfos));
	}
	
	//ZSS-1193, ZSS-1233: should cached in autofilter by column index
	//@since 3.9.0
	//@Internal
	//@See CustomFiltersCtrl
	public List<FilterRowInfo> getCachedSet(int index) {
		return this.orderedRowInfosMap.get(index);
	}
	
	//ZSS-1234
	//@since 3.9.0
	//@Internal
	//@See CustomFiltersCtrl
	public void setFilterType(int index, int type) {
		this.filterTypes.put(index, type);
	}
	
	//ZSS-1234
	//@since 3.9.0
	//@Internal
	//@See CustomFiltersCtrl
	public int getFilterType(int index) {
		return this.filterTypes.get(index);
	}	
}
