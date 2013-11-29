package org.zkoss.zss.ngmodel;

public interface NHyperlink {
	public enum HyperlinkType {
		URL, DOCUMENT, EMAIL, FILE
	}

	public HyperlinkType getType();

	public String getAddress();

	public String getLabel();
}