/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface NHyperlink {
	public enum HyperlinkType {
		URL, DOCUMENT, EMAIL, FILE
	}

	public HyperlinkType getType();

	public String getAddress();

	public String getLabel();
	
	public void setType(HyperlinkType type);
	
	public void setAddress(String address);
	
	public void setLabel(String label);
}