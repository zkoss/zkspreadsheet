/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2014/01/14 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl;

import java.util.Collection;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.impl.*;
import org.zkoss.zss.ngapi.NRange;

/**
 * A helper class that cooperate with NRange to manipulate pictures.
 * @author Hawk
 *
 */
class PictureHelper extends RangeHelperBase {

	public PictureHelper(NRange range) {
		super(range);
	}
	
	public Picture addPicture(ClientAnchor anchor, byte[] image, int format) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			final Picture picture = dm.addPicture(_sheet, anchor, image, format);
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyPictureAdd(ref, picture);
			}
			return picture;
		}
	}
	
	public void deletePicture(Picture picture) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			ClientAnchor anchor = picture.getPreferredSize();
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			// ZSS-397: keep picture ID for notifying
			// the picture data including ID will be gone after deleting
			final String id = picture.getPictureId();
			dm.deletePicture(_sheet, picture); //must after getPreferredSize() or anchor is gone!
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyPictureDelete(ref, id);
			}
		}
	}

	@Override
	public void movePicture(Picture picture, ClientAnchor anchor) {
		synchronized (_sheet) {
			DrawingManager dm = ((SheetCtrl)_sheet).getDrawingManager();
			dm.movePicture(_sheet, picture, anchor);
			final XRangeImpl rng = (XRangeImpl) XRanges.range(_sheet, anchor.getRow1(), anchor.getCol1(), anchor.getRow2(), anchor.getCol2());
			final Collection<Ref> refs = rng.getRefs();
			if (refs != null && !refs.isEmpty()) {
				final Ref ref = refs.iterator().next();
				BookHelper.notifyPictureUpdate(ref, picture);
			}
		}
	}

	
}
