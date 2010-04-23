/* CodeView.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Fri May 23 15:03:13     2008, Created by ivancheng
}}IS_NOTE

Copyright (C) 2008 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.demo;

import org.zkoss.zk.ui.*;
import org.zkoss.zk.ui.ext.AfterCompose;
import org.zkoss.zul.*;

/**
 * The code textbox.
 * 
 * @author ivancheng
 */
public class CodeView extends Textbox implements AfterCompose {
	public void afterCompose() {
		execute();
	}
	public void execute() {
		Component view = getFellow("view");
		Components.removeAllChildren(view);
		Executions.createComponentsDirectly(getValue(), "zul", view, null);
	}
}
