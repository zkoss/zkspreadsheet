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
import org.zkoss.zul.Tab;
import org.zkoss.zul.Textbox;

/**
 * @author Sam
 *
 */
public class MainLayout extends GenericForwardComposer {

	Tab app0; //ZSS Apreadsheet App
	Tab app0Src;
	Tab app2; //ZK Spreadsheet H1N1 Demo
	Tab app2Src;
	
	Div demo1Cave;
	Div demo2Cave;
	Textbox demo1CodeView;
	Textbox demo2CodeView;
	
	private boolean isCurrentApp0 = false;

	private void clearDemos() {
		Components.removeAllChildren(demo1Cave);
		Components.removeAllChildren(demo2Cave);
	}
	
	private void createDemo(String code, Component parent) {
		clearDemos();
		try {
			Executions.createComponentsDirectly(code, "zul", parent, null);
		} catch (RuntimeException e) {
			if ("true".equalsIgnoreCase(System.getProperty("zkdemo.debug")))
				System.out.println("\n Error caused by zkdemo at : " + new java.util.Date() + "\n code: " + code);
			throw e;
		}
	}
	
	public void onClick$app0() {
		if (!isCurrentApp0) {
			createDemo(demo1CodeView.getValue(), demo1Cave);
			isCurrentApp0 = true;
		}
	}
	
	public void onClick$app2() {
		if (isCurrentApp0) {
			createDemo(demo2CodeView.getValue(), demo2Cave);
			isCurrentApp0 = false;
		}
	}
	
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		createDemo(demo1CodeView.getValue(), demo1Cave);
		isCurrentApp0 = true;
	}
}