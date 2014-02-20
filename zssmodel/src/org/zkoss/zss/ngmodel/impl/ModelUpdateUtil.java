package org.zkoss.zss.ngmodel.impl;

import java.util.Set;
import org.zkoss.zss.ngapi.impl.CellUpdateCollector;
import org.zkoss.zss.ngapi.impl.InsertDeleteUpdateCollector;
import org.zkoss.zss.ngapi.impl.RefUpdateCollector;
import org.zkoss.zss.ngapi.impl.MergeUpdateCollector;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTable;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/*package*/ class ModelUpdateUtil {

	
	/*package*/ static void handlePrecedentUpdate(NBookSeries bookSeries, Ref precedent){
		//clear formula cache (that reval the unexisted sheet before
		FormulaCacheCleaner clearer = FormulaCacheCleaner.getCurrent();
		RefUpdateCollector collector = RefUpdateCollector.getCurrent();
		Set<Ref> dependents = null; 
		//get table when collector and clearer is not ignored (in import case, we should ignore clear cahche)
		if(collector!=null || clearer!=null || bookSeries.isAutoFormulaCacheClean()){
			DependencyTable table = ((AbstractBookSeriesAdv)bookSeries).getDependencyTable();
			dependents = table.getDependents(precedent); 
		}
		addRefUpdate(precedent);
		if(dependents!=null && dependents.size()>0){
			if(clearer!=null){
				clearer.clear(dependents);
			}else if(bookSeries.isAutoFormulaCacheClean()){
				new FormulaCacheClearHelper(bookSeries).clear(dependents);
			}
			if(collector!=null){
				collector.addRefs(dependents);
			}
		}
	}

	/*package*/ static void addRefUpdate(Ref ref) {
		RefUpdateCollector collector = RefUpdateCollector.getCurrent();
		if(collector!=null){
			collector.addRef(ref);
		}
	}
	
	/*package*/ static void addCellUpdate(NSheet sheet,int row,int column){
		addCellUpdate(sheet,row,column,row,column);
	}
	/*package*/ static void addCellUpdate(NSheet sheet,int row,int column, int lastRow, int lastColumn){
		CellUpdateCollector collector = CellUpdateCollector.getCurrent();
		if(collector!=null){
			collector.addCellUpdate(sheet,row, column,lastRow,lastColumn);
		}
	}
	
	/*package*/ static void addMergeUpdate(NSheet sheet,CellRegion original,CellRegion changeTo){
		MergeUpdateCollector collector = MergeUpdateCollector.getCurrent();
		if(collector!=null){
			collector.addMergeChange(sheet,original,changeTo);
		}
	}

	/*package*/static void addInsertDeleteUpdate(NSheet sheet, boolean inserted, boolean isRow, int index, int lastIndex) {
		InsertDeleteUpdateCollector collector = InsertDeleteUpdateCollector.getCurrent();
		if(collector != null) {
			collector.addInsertDeleteUpdate(sheet, inserted, isRow, index, lastIndex);
		}
	}
}
