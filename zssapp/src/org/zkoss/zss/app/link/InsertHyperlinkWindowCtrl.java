/* InsertHyperlinkWindowCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 7, 2010 11:34:37 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.link;

import java.util.Iterator;
import java.util.LinkedHashMap;

import org.zkoss.lang.Objects;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zul.Button;
import org.zkoss.zul.Include;
import org.zkoss.zul.Textbox;

/**
 * @author Sam
 *
 */
public class InsertHyperlinkWindowCtrl extends GenericForwardComposer {
	
	private Button webBtn;
	private Button docBtn;
	private Button mailBtn;
	private LinkedHashMap<Button, String> linkTypeBtns = new LinkedHashMap<Button, String>(3);
	
	private final static String WEBLINK_CONTENT_URI="/menus/hyperlink/webLink.zul";
	private final static String DOCLINK_CONTENT_URI="/menus/hyperlink/docLink.zul";
	private final static String MAILLINK_CONTENT_URI="/menus/hyperlink/mailLink.zul";
	private Include content;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		initButton();
	}
	/*  */
	private void initButton() {
		linkTypeBtns.put(webBtn, WEBLINK_CONTENT_URI);
		linkTypeBtns.put(docBtn, DOCLINK_CONTENT_URI);
		linkTypeBtns.put(mailBtn, MAILLINK_CONTENT_URI);
	}



	public void onClick$webBtn() {
		setLinkType(webBtn);
	}
	public void onClick$docBtn() {
		setLinkType(docBtn);
	}
	public void onClick$mailBtn() {
		setLinkType(mailBtn);
	}
	
	private void setLinkType(Button btn) {
		for (Iterator<Button> iter = linkTypeBtns.keySet().iterator(); iter.hasNext();) {
			Button b = iter.next();
			//set selected button disable
			b.setDisabled( Objects.equals(btn, b) );
		}
		content.setSrc( linkTypeBtns.get(btn) );
	}
	
	/**
	 * Returns the string of web page address, return null if component not found
	 * @return
	 */
	public String getWebAddress() {
		Textbox mailAddr = (Textbox)content.getFellowIfAny("mailAddr");
		return mailAddr != null ? mailAddr.getText() : "";
	}
	
	/**
	 * Returns the cell reference string, return null if component not found
	 * @return
	 */
	public String getCellRef() {
		Textbox cellRef = (Textbox)content.getFellowIfAny("cellRef");
		return cellRef != null ? cellRef.getText() : "";
	}
	
	/**
	 * Returns the reference sheet name, return null if link type is not document
	 * @return
	 */
	public String getRefSheet() {
		Textbox refSheet = (Textbox)content.getFellowIfAny("refSheet");
		return refSheet != null ? refSheet.getText() : "";	
	}
}