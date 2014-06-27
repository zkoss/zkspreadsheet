package org.zkoss.zss.range.impl;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import org.zkoss.util.logging.Log;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.sys.dependency.ObjectRef;
import org.zkoss.zss.model.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.dependency.Ref.RefType;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
/*package*/ class RefNotifyContentChangeHelper extends RefHelperBase{

	private static final Log _logger = Log.lookup(RefNotifyContentChangeHelper.class);
			
	private NotifyChangeHelper _notifyHelper = new NotifyChangeHelper();
	public RefNotifyContentChangeHelper(SBookSeries bookSeries) {
		super(bookSeries);
	}
	public void notifyContentChange(Ref notify) {
		if (notify.getType() == RefType.CELL || notify.getType() == RefType.AREA) {
			handleAreaRef(notify);
		} else if (notify.getType() == RefType.OBJECT) {
			if(((ObjectRef)notify).getObjectType()==ObjectType.CHART){
				handleChartRef((ObjectRef)notify);
			}else if(((ObjectRef)notify).getObjectType()==ObjectType.DATA_VALIDATION){
				handleDataValidationRef((ObjectRef)notify);
			}else if(((ObjectRef)notify).getObjectType()==ObjectType.AUTO_FILTER){ // ZSS-555
				handleAutoFilterRef((ObjectRef)notify);
			}
		} else {// TODO another

		}
	}
	public void notifyContentChange(Set<Ref> notifySet) {
		Map<String,Ref> chartDependents  = new LinkedHashMap<String, Ref>();
		Map<String,Ref> validationDependents  = new LinkedHashMap<String, Ref>();	
		
		// clear formula cache
		for (Ref notify : notifySet) {
			if(_logger.debugable()){
				_logger.debug("Notify Dependent Change : "+notify);
			}
			//clear the dependent's formula cache since the precedent is changed.
			if (notify.getType() == RefType.CELL || notify.getType() == RefType.AREA) {
				handleAreaRef(notify);
			} else if (notify.getType() == RefType.OBJECT) {
				if(((ObjectRef)notify).getObjectType()==ObjectType.CHART){
					chartDependents.put(((ObjectRef)notify).getObjectIdPath()[0], notify);
				}else if(((ObjectRef)notify).getObjectType()==ObjectType.DATA_VALIDATION){
					validationDependents.put(((ObjectRef)notify).getObjectIdPath()[0], notify);
				}else if(((ObjectRef)notify).getObjectType()==ObjectType.AUTO_FILTER){
					handleAutoFilterRef((ObjectRef)notify);
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
		SBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		String[] ids = notify.getObjectIdPath();
		_notifyHelper.notifyChartChange(sheet,ids[0]);
				
	}
	
	private void handleDataValidationRef(ObjectRef notify) {
		SBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		String[] ids = notify.getObjectIdPath();
		_notifyHelper.notifyDataValidationChange(sheet,ids[0]);
	}

	// ZSS-555
	private void handleAutoFilterRef(ObjectRef notify) {
		SBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		_notifyHelper.notifySheetAutoFilterChange(sheet);
	}

	private void handleAreaRef(Ref notify) {
		SBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		SSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		_notifyHelper.notifyCellChange(new SheetRegion(sheet,notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn()));
	}
}
