package org.zkoss.zss.model.impl;

import java.util.Collection;
import java.util.Collections;
import java.util.TreeMap;

import org.zkoss.zss.model.CellRegion;
/**
 * The auto fitler implement
 * @author dennis
 * @since 3.5.0
 */
public class AutoFilterImpl extends AbstractAutoFilterAdv {
	private static final long serialVersionUID = 1L;
	
	private final CellRegion _region;
	
	private final TreeMap<Integer,NFilterColumn> _columns;

	public AutoFilterImpl(CellRegion region){
		this._region = region;
		_columns = new TreeMap<Integer,NFilterColumn>();
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

}
