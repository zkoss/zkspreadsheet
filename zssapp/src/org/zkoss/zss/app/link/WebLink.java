/* WebLink.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 27, 2010 6:04:06 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.link;

import java.util.List;

import org.zkoss.zk.ui.HtmlMacroComponent;
import org.zkoss.zul.Combobox;

/**
 * @author Sam
 *
 */
public class WebLink extends HtmlMacroComponent {
	
	public WebLink() {
		setMacroURI("/menus/hyperlink/webLink.zu");
	}
	
	public void setAvailableUrls(List urls) {
		
	}
	
	public String getUrl() {
		Combobox addrCB = (Combobox)getFellow("addrCombobox");
		return addrCB.getText();
	}
}
