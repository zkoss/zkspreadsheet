package org.zkoss.zss.ngmodel.impl;

import java.util.Set;

import org.zkoss.zss.ngapi.impl.CellUpdateCollector;
import org.zkoss.zss.ngapi.impl.DependentUpdateCollector;
import org.zkoss.zss.ngapi.impl.MergeUpdateCollector;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/*package*/ class ModelUpdateUtil {

	
	/*package*/ static void handleDependentUpdate(NBookSeries bookSeries, Ref precedent){
		//clear formula cache (that reval the unexisted sheet before
		FormulaCacheCleaner clearer = FormulaCacheCleaner.getCurrent();
		DependentUpdateCollector collector = DependentUpdateCollector.getCurrent();
		Set<Ref> dependents = null; 
		//get tabl when collector and clearer is not ignored (in import case, we should ignore clear cahche)
		if(collector!=null || clearer!=null || bookSeries.isAutoFormulaCacheClean()){
			DependencyTable table = ((AbstractBookSeriesAdv)bookSeries).getDependencyTable();
			dependents = table.getDependents(precedent); 
		}
		if(dependents!=null && dependents.size()>0){
			if(clearer!=null){
				clearer.clear(dependents);
			}else if(bookSeries.isAutoFormulaCacheClean()){
				new FormulaCacheClearHelper(bookSeries).clear(dependents);
			}
			if(collector!=null){
				collector.addDependents(dependents);
			}
		}
	}
	
	/*package*/ static void addCellUpdate(int row,int column){
		CellUpdateCollector collector = CellUpdateCollector.getCurrent();
		if(collector!=null){
			collector.addCellUpdate(row, column);
		}
	}
	
	/*package*/ static void addMergeUpdate(CellRegion original,CellRegion changeTo){
		MergeUpdateCollector collector = MergeUpdateCollector.getCurrent();
		if(collector!=null){
			collector.addMergeChange(original,changeTo);
		}
	}
}
