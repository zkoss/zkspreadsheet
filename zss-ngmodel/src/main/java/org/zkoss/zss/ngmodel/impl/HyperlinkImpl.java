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
package org.zkoss.zss.ngmodel.impl;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class HyperlinkImpl extends AbstractHyperlinkAdv {

	private static final long serialVersionUID = 1L;
	
	private HyperlinkType type;
	private String address;
	private String label;
	
	public HyperlinkImpl(){}
	public HyperlinkImpl(HyperlinkType type,String address, String label){
		this.type = type;
		this.address = address;
		this.label = label;
	}
	
	public HyperlinkType getType() {
		return type;
	}
	public void setType(HyperlinkType type) {
		this.type = type;
	}
	public String getAddress() {
		return address;
	}
	public void setAddress(String address) {
		this.address = address;
	}
	public String getLabel() {
		return label;
	}
	public void setLabel(String label) {
		this.label = label;
	}
	@Override
	public AbstractHyperlinkAdv clone() {
		return new HyperlinkImpl(type,address,label);
	}
}
