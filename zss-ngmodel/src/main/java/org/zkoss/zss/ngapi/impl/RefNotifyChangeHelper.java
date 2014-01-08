package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.AbstractBookAdv;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

/*package*/ class RefNotifyChangeHelper {
	final NBookSeries bookSeries;
	public RefNotifyChangeHelper(NBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	public void notifySizeChange(HashSet<Ref> notifySet) {
		for (Ref notify : notifySet) {
			System.out.println(">>> Notify Size "+notify);
			if (notify.getType() == RefType.CELL || notify.getType() == RefType.AREA) {
				handleRowColumnSizeChange(notify);
			} 
		}
	}

	private void handleRowColumnSizeChange(Ref notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_ROW_COLUMN_SIZE_CHANGE,sheet,
				new CellRegion(notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn())));
	}
	
	public void notifySheetAutoFilterChange(Ref notify) {
		if (notify.getType() != RefType.SHEET) {
			return;
		}
		System.out.println(">>> Notify Autofilter "+notify);
		NBook book = bookSeries.getBook(notify.getBookName());
		if(book==null) return;
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		if(sheet==null) return;
		((AbstractBookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_AUTOFILTER_CHANGE,sheet));
	}
}
