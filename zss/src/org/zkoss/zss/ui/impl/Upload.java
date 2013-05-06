/* UploadGhost.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 5, 2012 7:06:09 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import org.zkoss.zk.ui.AbstractComponent;
import org.zkoss.zss.ui.sys.XActionHandler;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;

/**
 * Internal Use Only. for {@link XActionHandler}.
 * 
 * @author sam
 *
 */
public class Upload extends AbstractComponent {
	
	public Upload() {
		this.setAttribute(SpreadsheetCtrl.CHILD_PASSING_KEY, true);
	}
}
