/* ReloadButton.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Fri May 23 16:38:25     2008, Created by ivancheng
}}IS_NOTE

Copyright (C) 2008 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.demo;

import java.util.Set;
import java.util.HashSet;

import org.zkoss.zk.ui.*;
import org.zkoss.zul.*;

/**
 * The "Reload" button.
 * 
 * @author ivancheng
 */
public class ReloadButton extends Button {
	public void onClick() {
		Path.getComponent("//zssUserGuide/showRoom/contentArea").invalidate();
	}
}
