package org.zkoss.zss.ngmodel.impl;

import java.util.Set;

import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/**
 * 
 * @author dennis
 *
 */
public class FormulaCacheClearer {

	static ThreadLocal<FormulaCacheClearer>  current = new ThreadLocal<FormulaCacheClearer>();
	
	final private NBookSeries bookSeries;
	final private boolean ignoreClear;
	
	public FormulaCacheClearer(){
		this.bookSeries = null;
		this.ignoreClear = true;
	}
	public FormulaCacheClearer(NBookSeries bookSeries){
		this.bookSeries = bookSeries;
		this.ignoreClear = false;
	}

	public static FormulaCacheClearer setCurrent(FormulaCacheClearer ctx){
		FormulaCacheClearer old = current.get();
		current.set(ctx);
		return old;
	}
	
	public static FormulaCacheClearer getCurrent(){
		return current.get();
	}
	
	public void clear(Set<Ref> dependents){
		if(ignoreClear)
			return;
		new FormulaCacheClearHelper(bookSeries).clear(dependents);
	}
	
	public boolean isIgnore(){
		return ignoreClear;
	}

	public void clearByPrecedent(Ref precedent) {
		if(ignoreClear)
			return;
		DependencyTable table = ((AbstractBookSeriesAdv)bookSeries).getDependencyTable();
		Set<Ref> dependents = table.getDependents(precedent);
		if(dependents.size()>0){
			clear(dependents);
		}
	}
}
