/* HyperlinkImpl.java

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

import org.zkoss.zss.api.model.Hyperlink;

/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class HyperlinkImpl implements Hyperlink{

	ModelRef<org.zkoss.poi.ss.usermodel.Hyperlink> linkRef;

	public HyperlinkImpl(ModelRef<org.zkoss.poi.ss.usermodel.Hyperlink> linkRef) {
		this.linkRef = linkRef;
	}
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((linkRef == null) ? 0 : linkRef.hashCode());
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
		if (linkRef == null) {
			if (other.colorRef != null)
				return false;
		} else if (!linkRef.equals(other.colorRef))
			return false;
		return true;
	}

	public org.zkoss.poi.ss.usermodel.Hyperlink getNative() {
		return linkRef.get();
	}
	
	@Override
	public HyperlinkType getType() {
		return EnumUtil.toHyperlinkType(getNative().getType());
	}

	@Override
	public String getAddress() {
		return getNative().getAddress();
	}

	@Override
	public String getLabel() {
		return getNative().getLabel();
	}
}
