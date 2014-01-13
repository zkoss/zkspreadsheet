package org.zkoss.zss.ngapi.impl;

import java.util.HashSet;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.SheetRegion;
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
	
	
	public void notifySheetAutoFilterChange(SheetRegion notify) {
		((AbstractBookAdv) notify.getSheet().getBook())
				.sendModelEvent(ModelEvents.createModelEvent(
						ModelEvents.ON_AUTOFILTER_CHANGE, notify.getSheet()));
	}

	public void notifySheetFreezeChange(SheetRegion notify) {
		((AbstractBookAdv) notify.getSheet().getBook())
				.sendModelEvent(ModelEvents.createModelEvent(
						ModelEvents.ON_FREEZE_CHANGE, notify.getSheet()));
	}
	
}
