/* DemoWindowComposer.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 6, 2010 12:15:31 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.demo;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zul.Button;
import org.zkoss.zul.Tab;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

/**
 * @author Sam
 *
 */
public class DemoWindowComposer extends GenericForwardComposer {
	Window view;
	Tab demoView;
	Textbox codeView;
	Button reloadBtn;
	Button tryBtn;
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		if (view != null) execute();
	}
	public void execute() {
		Components.removeAllChildren(view);
		String code = codeView.getValue();
		try {
			Executions.createComponentsDirectly(code, "zul", view, null);
		} catch (RuntimeException e) {
			if ("true".equalsIgnoreCase(System.getProperty("zkdemo.debug")))
				System.out.println("\n Error caused by zkdemo at : " + new java.util.Date() + "\n code: " + code);
			throw e;
		}
	}
	public void onClick$reloadBtn(Event event) {
		demoView.setSelected(true);
	}
	public void onClick$tryBtn(Event event) {
		demoView.setSelected(true);
		execute();
	}
}