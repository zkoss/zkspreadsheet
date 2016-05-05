/* ExtraFillImpl.java

	Purpose:
		
	Description:
		
	History:
		Feb 1, 2016 2:23:48 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SFill;
import org.zkoss.zss.model.SBook;

//ZSS-1162
/**
 * Fill in ExtraStyle is a bit different from in  CellStyle.
 * 
 * @author henri
 * @sine 3.8.3
 */
public class ExtraFillImpl extends FillImpl {
	private static final long serialVersionUID = -6167978135008154858L;

	//ZSS-1183
	//@since 3.9.0
	/*package*/ ExtraFillImpl(ExtraFillImpl src, SBook book) {
		super(src, book);
	}
	
	public ExtraFillImpl() {
		super();
	}
	public ExtraFillImpl(FillPattern pattern, String fgColor, String bgColor) {
		super(pattern, fgColor, bgColor);
	}
	//ZSS-1140
	public ExtraFillImpl(FillPattern pattern, SColor fgColor, SColor bgColor) {
		super(pattern, fgColor, bgColor);
	}

	// FillPattern in Dxf(ExtraStyle) is default to FillPattern.SOLID, not like in 
	// xf(CellStyle) which is FillPattern.NONE.
	@Override
	public FillPattern getFillPattern() {
		return _fillPattern == null ? FillPattern.SOLID : _fillPattern; //ZSS-1145
	}
	
	//ZSS-1183
	@Override
	/*package*/ SFill cloneFill(SBook book) {
		return book == null ? this : new ExtraFillImpl(this, book);
	}
}
