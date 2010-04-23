package org.zkoss.zss.app;

import org.zkoss.lang.Library;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Components;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.util.GenericForwardComposer;
import org.zkoss.zul.Button;
import org.zkoss.zul.Div;
import org.zkoss.zul.Image;
import org.zkoss.zul.Tab;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

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
		//System.out.println(code);
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
		Path.getComponent("//zssapp/xcontents").invalidate();
	}
	public void onClick$tryBtn(Event event) {
		demoView.setSelected(true);
		execute();
	}
}
