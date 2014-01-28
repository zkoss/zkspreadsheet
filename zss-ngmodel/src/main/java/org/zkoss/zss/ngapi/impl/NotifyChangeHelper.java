package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.impl.AbstractBookAdv;

/*package*/ class NotifyChangeHelper{

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
				sheet, ModelEvents.createDataMap(ModelEvents.PARAM_OBJECT_ID, picture.getId())));
	}

	public void notifySheetPictureDelete(NSheet sheet, String id) {
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_PICTURE_DELETE, 
				sheet, ModelEvents.createDataMap(ModelEvents.PARAM_OBJECT_ID, id)));
	}

	public void notifySheetPictureMove(NSheet sheet, NPicture picture) {
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_PICTURE_UPDATE, 
				sheet, ModelEvents.createDataMap(ModelEvents.PARAM_OBJECT_ID, picture.getId())));
	}

	public void notifyMergeChange(Set<MergeUpdate> mergeNotifySet) {
		LinkedHashSet<SheetRegion> toRemove = new LinkedHashSet<SheetRegion>();
		LinkedHashSet<SheetRegion> toAdd = new LinkedHashSet<SheetRegion>();
		for(MergeUpdate mu:mergeNotifySet){
			SheetRegion remove = mu.getOrgMerge()==null?null:new SheetRegion(mu.getSheet(),mu.getOrgMerge());
			SheetRegion add = mu.getMerge()==null?null:new SheetRegion(mu.getSheet(),mu.getMerge());
			if(remove!=null){
				toRemove.add(remove);
				toAdd.remove(remove);
			}
			if(add!=null){
				toAdd.add(add);
				toRemove.remove(add);
			}
		}
		for(SheetRegion notify:toRemove){//remove the final remove list
			NBook book = notify.getSheet().getBook();
			System.out.println(">>> Notify remove merge "+notify.getRegion().getReferenceString());
			((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_MERGE_DELETE,notify.getSheet(),
					notify.getRegion()));
			
		}
		for(SheetRegion notify:toAdd){
			NBook book = notify.getSheet().getBook();
			System.out.println(">>> Notify add merge "+notify.getRegion().getReferenceString());
			((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_MERGE_ADD,notify.getSheet(),
					notify.getRegion()));
		}
		
	}

	public void notifyCellChange(Set<SheetRegion> cellNotifySet) {
		for(SheetRegion notify:cellNotifySet){
			NBook book = notify.getSheet().getBook();
			System.out.println(">>> Notify update cell "+notify.getRegion().getReferenceString());
			((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_CELL_CONTENT_CHANGE,notify.getSheet(),
				new CellRegion(notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn())));
		}
	}
	
	public void notifySheetDelete(NBook book,NSheet deletedSheet,int deletedIndex){
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_DELETE,book,
				ModelEvents.createDataMap(ModelEvents.PARAM_SHEET,deletedSheet,ModelEvents.PARAM_INDEX,deletedIndex)));
	}
	
	public void notifySheetCreate(NSheet sheet){
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_CREATE,sheet));
	}
	
	public void notifySheetNameChange(NSheet sheet,String oldName){
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_NAME_CHANGE,sheet,
				ModelEvents.createDataMap(ModelEvents.PARAM_OLD_NAME,oldName)));
	}
	
	public void notifySheetReorder(NSheet sheet,int oldIdx){
		((AbstractBookAdv) sheet.getBook()).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_SHEET_ORDER_CHANGE,sheet,
				ModelEvents.createDataMap(ModelEvents.PARAM_OLD_INDEX,oldIdx)));
	}
	
}
