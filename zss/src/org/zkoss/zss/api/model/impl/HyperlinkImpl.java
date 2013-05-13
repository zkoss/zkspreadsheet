package org.zkoss.zss.api.model.impl;

import org.zkoss.zss.api.model.Hyperlink;
import org.zkoss.zss.model.sys.XBook;

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
