package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.SheetRegion;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

/*package*/ class RefNotifyContentChangeHelper extends RefHelperBase{

	NotifyChangeHelper notifyHelper = new NotifyChangeHelper();
	public RefNotifyContentChangeHelper(NBookSeries bookSeries) {
		super(bookSeries);
	}

	public void notifyContentChange(HashSet<Ref> notifySet) {
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();	
		
		// clear formula cache
		for (Ref notify : notifySet) {
			System.out.println(">>> Notify Dependent Change : "+notify);
			//clear the dependent's formula cache since the precedent is changed.
			if (notify.getType() == RefType.CELL || notify.getType() == RefType.AREA) {
				handleCellRef(notify);
			} else if (notify.getType() == RefType.OBJECT) {
				if(((ObjectRef)notify).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)notify).getObjectIdPath()[0], notify);
				}else if(((ObjectRef)notify).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)notify).getObjectIdPath()[0], notify);
				}
			} else {// TODO another

			}
		}
		
		for (Ref notify : chartDependents.values()) {
			handleChartRef((ObjectRef)notify);
		}
		for (Ref notify : validationDependents.values()) {
			handleDataValidationRef((ObjectRef)notify);
		}	
	}

	private void handleChartRef(ObjectRef notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		String[] ids = notify.getObjectIdPath();
		notifyHelper.notifyChartChange(sheet,ids[0]);
				
	}
	
	private void handleDataValidationRef(ObjectRef notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		String[] ids = notify.getObjectIdPath();
		notifyHelper.notifyDataValidationChange(sheet,ids[0]);
	}

	private void handleCellRef(Ref notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		notifyHelper.notifyCellChange(new SheetRegion(sheet,notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn()));
	}
}
