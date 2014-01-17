package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.impl.AbstractBookAdv;

/*package*/ class NotifyChangeHelper extends RangeHelperBase{
	public NotifyChangeHelper(NRange range) {
		super(range);
	}

	public void notifySizeChange(HashSet<SheetRegion> notifySet) {
		for (SheetRegion notify : notifySet) {
			notifySizeChange(notify);
		}
	}

	public void notifySizeChange(SheetRegion notify) {
		((AbstractBookAdv) notify.getSheet().getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_ROW_COLUMN_SIZE_CHANGE,
				notify.getSheet(),
				new CellRegion(notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn())));
	}
	
	
	public void notifySheetAutoFilterChange(NSheet sheet) {
		((AbstractBookAdv) sheet.getBook())
				.sendModelEvent(ModelEvents.createModelEvent(
						ModelEvents.ON_AUTOFILTER_CHANGE, sheet));
	}

	public void notifySheetFreezeChange(NSheet sheet) {
		((AbstractBookAdv) sheet.getBook())
				.sendModelEvent(ModelEvents.createModelEvent(
						ModelEvents.ON_FREEZE_CHANGE, sheet));
	}
	
	public void notifySheetPictureAdd(NSheet sheet, NPicture picture){
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_PICTURE_ADD, 
				sheet, ModelEvents.createDataMap(ModelEvents.PARAM_PICTURE, picture)));
	}

	public void notifySheetPictureDelete(NSheet sheet, NPicture picture) {
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_PICTURE_DELETE, 
				sheet, ModelEvents.createDataMap(ModelEvents.PARAM_PICTURE, picture)));
	}

	public void notifySheetPictureMove(NSheet sheet, NPicture picture) {
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_PICTURE_UPDATE, 
				sheet, ModelEvents.createDataMap(ModelEvents.PARAM_PICTURE, picture)));
	}

	public void notifyMergeChange(NSheet sheet,Set<MergeUpdate> mergeNotifySet) {
		LinkedHashSet<CellRegion> toRemove = new LinkedHashSet<CellRegion>();
		LinkedHashSet<CellRegion> toAdd = new LinkedHashSet<CellRegion>();
		for(MergeUpdate mu:mergeNotifySet){
			CellRegion remove = mu.getOrgMerge();
			CellRegion add = mu.getMerge();
			if(remove!=null){
				toRemove.add(remove);
				toAdd.remove(remove);
			}
			if(add!=null){
				toAdd.add(add);
				toRemove.remove(add);
			}
		}
		NBook book = sheet.getBook();
		for(CellRegion notify:toRemove){//remove the final remove list
			System.out.println(">>> notify remove merge "+notify.getReferenceString());
			((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_MERGE_DELETE,sheet,
					notify));
			
		}
		for(CellRegion notify:toAdd){
			System.out.println(">>> notify add merge "+notify.getReferenceString());
			((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_MERGE_ADD,sheet,
					notify));
		}
		
	}

	public void notifyCellChange(NSheet sheet,Set<CellRegion> cellNotifySet) {
		NBook book = sheet.getBook();
		for(CellRegion notify:cellNotifySet){
			System.out.println(">>> notify update cell "+notify.getReferenceString());
			((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_CELL_CONTENT_CHANGE,sheet,
				new CellRegion(notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn())));
		}
	}
	
}
