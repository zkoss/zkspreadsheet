package org.zkoss.zss.demo;

import java.util.Date;
import java.util.Map;

import org.zkoss.zk.ui.Page;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.util.Initiator;

public class PageInit implements Initiator {

	static Date startTime = new Date();
	
	
	public void doInit(Page page, Map<String, Object> arg1) throws Exception {
		page.setAttribute("timestamp", startTime.getTime());
		page.setAttribute("starttime", startTime.getTime());
		page.setAttribute("buildVer", WebApps.getCurrent().getBuild());
	}
}
