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

	ModelRef<XBook> bookRef;
	ModelRef<org.zkoss.poi.ss.usermodel.Color> colorRef;

	public ColorImpl(ModelRef<XBook> book, ModelRef<org.zkoss.poi.ss.usermodel.Color> color) {
		this.bookRef = book;
		this.colorRef = color;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((colorRef == null) ? 0 : colorRef.hashCode());
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
		if (colorRef == null) {
			if (other.colorRef != null)
				return false;
		} else if (!colorRef.equals(other.colorRef))
			return false;
		return true;
	}

	public org.zkoss.poi.ss.usermodel.Color getNative() {
		return colorRef.get();
	}
	public ModelRef<org.zkoss.poi.ss.usermodel.Color> getRef(){
		return colorRef;
	}

	public String getHtmlColor() {
		return BookHelper.colorToHTML(bookRef.get(),colorRef.get());
	}

}
