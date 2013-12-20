package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.BookAdv;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.dependency.Ref.RefType;

/*package*/ class RefNotifySizeChangeHelper {
	final NBookSeries bookSeries;
	public RefNotifySizeChangeHelper(NBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	public void handle(HashSet<Ref> dependentSet) {
		// clear formula cache
		for (Ref dependent : dependentSet) {
			System.out.println(">>> Notify Size"+dependent);
			//clear the dependent's formula cache since the precedent is changed.
			if (dependent.getType() == RefType.CELL || dependent.getType() == RefType.AREA) {
				handleCellRef(dependent);
			} else {// TODO another

			}
		}
	}


	private void handleCellRef(Ref notify) {
		NBook book = bookSeries.getBook(notify.getBookName());
		NSheet sheet = book.getSheetByName(notify.getSheetName());
		((BookAdv) book).sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_ROW_COLUMN_SIZE_CHANGE,sheet,
				new CellRegion(notify.getRow(),notify.getColumn(),notify.getLastRow(),notify.getLastColumn())));
	}
}
