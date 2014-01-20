package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.AbstractBookAdv;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef;
import org.zkoss.zss.ngmodel.sys.dependency.ObjectRef.ObjectType;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

/*package*/ class RefNotifyContentChangeHelper extends RefHelperBase{

	public RefNotifyContentChangeHelper(NBookSeries bookSeries) {
		super(bookSeries);
	}

	public void notifyContentChange(HashSet<Ref> notifySet) {
		// clear formula cache
		for (Ref notify : notifySet) {
			System.out.println(">>> Notify Dependent Change : "+notify);
			//clear the dependent's formula cache since the precedent is changed.
			if (notify.getType() == RefType.CELL || notify.getType() == RefType.AREA) {
				handleCellRef(notify);
			} else if (notify.getType() == RefType.OBJECT) {
				if(((ObjectRef)notify).getObjectType()==ObjectType.CHART){
					handleChartRef((ObjectRef)notify);
				}else if(((ObjectRef)notify).getObjectType()==ObjectType.DATA_VALIDATION){
					handleDataValidationRef((ObjectRef)notify);
				}
			} else {// TODO another

			}
		}
	}

	private void handleChartRef(ObjectRef notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		String[] ids = notify.getObjectIdPath();
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_CHART_CONTENT_CHANGE,sheet,
				ModelEvents.createDataMap(ModelEvents.PARAM_OBJECT_ID,ids[0])));
				
	}
	
	private void handleDataValidationRef(ObjectRef notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		String[] ids = notify.getObjectIdPath();
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_DATA_VALIDATION_CONTENT_CHANGE,sheet,
				ModelEvents.createDataMap(ModelEvents.PARAM_OBJECT_ID,ids[0])));
	}

	private void handleCellRef(Ref notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_CELL_CONTENT_CHANGE,sheet,
				new CellRegion(notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn())));
	}
}
