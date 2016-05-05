/* TableStyleElemImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 31, 2015 7:59:07 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBorder;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.STableStyleElem;

/**
 * @author henri
 * @since 3.8.0
 */
public class TableStyleElemImpl extends ExtraStyleImpl implements STableStyleElem {
	private static final long serialVersionUID = -8901383802206100226L;

	//ZSS-1183
	//@since 3.9.0
	/*package*/ TableStyleElemImpl(TableStyleElemImpl src, BookImpl book) {
		super(src, book);
	}
	

	public TableStyleElemImpl(SFont font, SFill fill, SBorder border) {
		super((AbstractFontAdv)font, (AbstractFillAdv) fill, (AbstractBorderAdv) border, null);
	}
	
	//ZSS-1183
	//@since 3.9.0
	@Override
	/*package*/ SCellStyle createCellStyle(SBook book) {
		return new TableStyleElemImpl(this, (BookImpl) book);
	}
}
