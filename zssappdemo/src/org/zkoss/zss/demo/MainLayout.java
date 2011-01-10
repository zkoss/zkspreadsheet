/* MainLayout.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 6, 2010 12:27:54 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.demo;


import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zul.Div;
import org.zkoss.zul.Include;

/**
 * @author Sam
 *
 */
public class MainLayout extends GenericForwardComposer {
	//TODO: use include will cause fire ON_WORKBOOK_OPEN multiple event
	// change to use div can fixed this issue, find out why
	//Include xcontents;
	Div xcontents;
	
	Div app0;
	Div app2;

	public void onClick$app0() {
		//zss live demo
		//xcontents.setSrc("app0.zul");
		
		Components.removeAllChildren(xcontents);
		Executions.createComponents("app0.zul", xcontents, null);
	}
	
	public void onClick$app2() {
		//gmap demo
		//xcontents.setSrc("app2.zul");
		
		Components.removeAllChildren(xcontents);
		Executions.createComponents("app2.zul", xcontents, null);
	}
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		//xcontents.setSrc("app0.zul");
		Executions.createComponents("app0.zul", xcontents, null);
	}
}