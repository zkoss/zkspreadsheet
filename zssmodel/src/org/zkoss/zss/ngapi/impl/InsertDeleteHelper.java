/* InsertDeleteHelper.java

	Purpose:
		
	Description:
		
	History:
		Feb 18, 2014 Created by Pao Wang

Copyright (C) 2014 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngapi.NRange;
import org.zkoss.zss.ngapi.NRange.DeleteShift;
import org.zkoss.zss.ngapi.NRange.InsertCopyOrigin;
import org.zkoss.zss.ngapi.NRange.InsertShift;

/**
 * A helper to perform insert/delete row/column/cells.
 * @author Pao
 */
public class InsertDeleteHelper extends RangeHelperBase {

	public InsertDeleteHelper(NRange range) {
		super(range);
	}

	public void insert(InsertShift shift, InsertCopyOrigin copyOrigin) {
		// just process on the first sheet even this range over multiple sheets

		// insert row/column/cell
		if(isWholeRow()) { // ignore insert direction
			sheet.insertRow(getRow(), getLastRow());

		} else if(isWholeColumn()) { // ignore insert direction
			sheet.insertColumn(getColumn(), getLastColumn());

		} else if(shift != InsertShift.DEFAULT) { // do nothing if "DEFAULT", it's according to XRange.insert() spec.
			sheet.insertCell(getRow(), getColumn(), getLastRow(), getLastColumn(), shift == InsertShift.RIGHT);
		}
		// TODO copy formal/style >>> in SheetImpl

	}

	public void delete(DeleteShift shift) {
		// just process on the first sheet even this range over multiple sheets

		// insert row/column/cell
		if(isWholeRow()) { // ignore insert direction
			sheet.deleteRow(getRow(), getLastRow());

		} else if(isWholeColumn()) { // ignore insert direction
			sheet.deleteColumn(getColumn(), getLastColumn());

		} else if(shift != DeleteShift.DEFAULT) { // do nothing if "DEFAULT", it's according to XRange.delete() spec.
			sheet.deleteCell(getRow(), getColumn(), getLastRow(), getLastColumn(), shift == DeleteShift.LEFT);
		}
	}
}
