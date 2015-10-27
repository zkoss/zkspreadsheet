/* ExtraStyleImpl.java

	Purpose:
		
	Description:
		
	History:
		Oct 26, 2015 3:58:03 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SBorder;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SFont;

/**
 * @author henri
 * @since 3.8.2
 */
public class ExtraStyleImpl extends CellStyleImpl implements SExtraStyle {
	private static final long serialVersionUID = -3304797955338410853L;

	public ExtraStyleImpl(SFont font, SFill fill, SBorder border, String dataFormat) {
		super((AbstractFontAdv)font, (AbstractFillAdv) fill, (AbstractBorderAdv) border);
		setDataFormat(dataFormat);
	}
}
