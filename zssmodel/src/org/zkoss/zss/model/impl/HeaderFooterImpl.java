package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SFooter;
import org.zkoss.zss.model.SHeader;



public class HeaderFooterImpl implements SHeader,SFooter,Serializable{

	private static final long serialVersionUID = 1L;
	
	String leftText = "";
	String rightText = "";
	String centerText = "";
	public String getLeftText() {
		return leftText;
	}
	public void setLeftText(String leftText) {
		this.leftText = leftText;
	}
	public String getRightText() {
		return rightText;
	}
	public void setRightText(String rightText) {
		this.rightText = rightText;
	}
	public String getCenterText() {
		return centerText;
	}
	public void setCenterText(String centerText) {
		this.centerText = centerText;
	}
}
