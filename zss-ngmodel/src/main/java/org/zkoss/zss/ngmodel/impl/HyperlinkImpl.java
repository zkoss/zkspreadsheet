package org.zkoss.zss.ngmodel.impl;

public class HyperlinkImpl extends HyperlinkAdv {

	private static final long serialVersionUID = 1L;
	
	private HyperlinkType type;
	private String address;
	private String label;
	
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
}
