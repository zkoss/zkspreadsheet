package org.zkoss.zss.ngmodel.impl;

import java.util.Set;

import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/*package*/ class DependentUpdateUtil {

	
	/*package*/ static void handleDependentUpdate(NBookSeries bookSeries, Ref precedent){
		//clear formula cache (that reval the unexisted sheet before
		FormulaCacheClearer clearer = FormulaCacheClearer.getCurrent();
		DependentCollector collector = DependentCollector.getCurrent();
		Set<Ref> dependents = null; 
		//get tabl when collector and clearer is not ignored (in import case, we should ignore clear cahche)
		if(collector!=null || (clearer!=null && !clearer.isIgnore()) || clearer==null){
			DependencyTable table = ((AbstractBookSeriesAdv)bookSeries).getDependencyTable();
			dependents = table.getDependents(precedent); 
		}
		if(dependents!=null && dependents.size()>0){
			if(clearer==null){
				//by default, if no clear, then we use default clear
				new FormulaCacheClearHelper(bookSeries).clear(dependents);
			}else if(clearer!=null && !clearer.isIgnore()){
				clearer.clear(dependents);
			}
			if(collector!=null){
				collector.addDependents(dependents);
			}
		}
	}
}
