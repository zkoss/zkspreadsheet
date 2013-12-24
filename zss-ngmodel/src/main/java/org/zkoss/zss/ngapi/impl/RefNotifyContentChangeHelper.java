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

/*package*/ class RefNotifyContentChangeHelper {
	final NBookSeries bookSeries;
	public RefNotifyContentChangeHelper(NBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	public void handle(HashSet<Ref> dependentSet) {
		// clear formula cache
		for (Ref dependent : dependentSet) {
			System.out.println(">>> Notify "+dependent);
			//clear the dependent's formula cache since the precedent is changed.
			if (dependent.getType() == RefType.CELL || dependent.getType() == RefType.AREA) {
				handleCellRef(dependent);
			} else if (dependent.getType() == RefType.OBJECT) {
				if(((ObjectRef)dependent).getObjectType()==ObjectType.CHART){
					handleChartRef((ObjectRef)dependent);
				}else if(((ObjectRef)dependent).getObjectType()==ObjectType.VALIDATION){
					
				}
			} else {// TODO another

			}
		}
	}

	private void handleChartRef(ObjectRef dependent) {
		NBook book = bookSeries.getBook(dependent
				.getBookName());
		NSheet sheet = book.getSheetByName(dependent
				.getSheetName());
		String[] ids = dependent.getObjectIdPath();
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_CHART_CONTENT_CHANGE,sheet,
				ModelEvents.PARAM_OBJECT_ID,ids[0]));
				
	}

	private void handleCellRef(Ref notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_CELL_CONTENT_CHANGE,sheet,
				new CellRegion(notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn())));
	}
}
