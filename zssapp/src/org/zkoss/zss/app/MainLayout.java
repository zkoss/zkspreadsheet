package org.zkoss.zss.app;

import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.ext.AfterCompose;
import org.zkoss.zul.Div;
import org.zkoss.zul.Include;
import org.zkoss.zul.Window;

public class MainLayout extends Window implements AfterCompose{
	Div _selected;
	Include xcontents;
	
	
	public void onCategorySelect(ForwardEvent event) {
		Div div = (Div) event.getOrigin().getTarget();
		String name=div.getId();
		if (_selected != div) {
			_selected = div;
			if(name.equals("app0"))
				xcontents.setSrc("onlineApp.zul");
			if(name.equals("app1"))
				xcontents.setSrc("app1.zul");
			if(name.equals("app2"))
				xcontents.setSrc("app2.zul");
		} 
	}


	
	public void afterCompose() {
		// TODO Auto-generated method stub
		xcontents=(Include) getFellow("xcontents");
		xcontents.setSrc("app2.zul");
	}
	
}