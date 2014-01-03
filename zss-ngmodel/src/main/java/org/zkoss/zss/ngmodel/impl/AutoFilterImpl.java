package org.zkoss.zss.ngmodel.impl;

import java.util.Collection;
import java.util.Collections;
import java.util.TreeMap;

import org.zkoss.zss.ngmodel.CellRegion;
/**
 * The auto fitler implement
 * @author dennis
 *
 */
public class AutoFilterImpl extends AbstractAutoFilterAdv {
	private static final long serialVersionUID = 1L;
	
	private final CellRegion region;
	
	private final TreeMap<Integer,NFilterColumn> columns;

	public AutoFilterImpl(CellRegion region){
		this.region = region;
		columns = new TreeMap<Integer,NFilterColumn>();
	}
	
	@Override
	public CellRegion getRegion() {
		return region;
	}

	@Override
	public Collection<NFilterColumn> getFilterColumns() {
		return Collections.unmodifiableCollection(columns.values());
	}

	@Override
	public NFilterColumn getFilterColumn(int index, boolean create) {
		NFilterColumn col = columns.get(index);
		if(col==null && create){
			int s = region.getLastColumn()-region.getColumn()+1; 
			if(index>=s){
				throw new IllegalStateException("the column index "+index+" >= "+s);
			}
			columns.put(index, col=new FilterColumnImpl(index));
		}
		return col;
	}

	@Override
	public void clearFilterColumn(int index) {
		columns.remove(index);
	}
	
	@Override
	public void clearFilterColumns() {
		columns.clear();
	}

}
