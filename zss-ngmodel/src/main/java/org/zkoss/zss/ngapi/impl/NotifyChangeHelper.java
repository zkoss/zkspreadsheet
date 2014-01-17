package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;
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

	public void notifyMergeChange(Set<MergeUpdate> mergeNotifySet) {
		// TODO Auto-generated method stub
		
	}

	public void notifyCellChange(Set<CellRegion> cellNotifySet) {
		// TODO Auto-generated method stub
		
	}
	
}
