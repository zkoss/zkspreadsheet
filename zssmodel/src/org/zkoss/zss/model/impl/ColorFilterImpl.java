/* ColorFilterImpl.java

	Purpose:
		
	Description:
		
	History:
		May 5, 2016 7:30:40 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SColorFilter;
import org.zkoss.zss.model.SExtraStyle;

//ZSS-1191
/**
 * Color filter associated with SAutoFilter
 * @author henri
 * @since 3.9.0
 */
public class ColorFilterImpl implements SColorFilter, Serializable {
	private boolean _byFontColor;
	private SExtraStyle _extraStyle;
	
	public ColorFilterImpl(SExtraStyle extraStyle, boolean byFontColor) {
		_extraStyle = extraStyle;
		_byFontColor = byFontColor;
	}
	
	@Override
	public SExtraStyle getExtraStyle() {
		return _extraStyle; 
	}

	@Override
	public boolean isByFontColor() {
		return _byFontColor;
	}
}
