/* ColorImpl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.Color;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class ColorImpl implements Color{

	private ModelRef<XBook> _bookRef;
	private ModelRef<org.zkoss.poi.ss.usermodel.Color> _colorRef;

	public ColorImpl(ModelRef<XBook> book, ModelRef<org.zkoss.poi.ss.usermodel.Color> color) {
		this._bookRef = book;
		this._colorRef = color;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((_colorRef == null) ? 0 : _colorRef.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		ColorImpl other = (ColorImpl) obj;
		if (_colorRef == null) {
			if (other._colorRef != null)
				return false;
		} else if (!_colorRef.equals(other._colorRef))
			return false;
		return true;
	}

	public org.zkoss.poi.ss.usermodel.Color getNative() {
		return _colorRef.get();
	}
	public ModelRef<org.zkoss.poi.ss.usermodel.Color> getRef(){
		return _colorRef;
	}

	public String getHtmlColor() {
		return BookHelper.colorToHTML(_bookRef.get(),_colorRef.get());
	}

}
