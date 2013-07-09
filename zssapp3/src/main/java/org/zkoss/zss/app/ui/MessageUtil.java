package org.zkoss.zss.app.ui;

import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Messagebox;

public class MessageUtil {
	
	public static void showInfoMessage(String message) {
		
		Clients.showNotification(message,"info",null,null,3000,true);
	}
	
	public static void showWarnMessage(String message) {
		Clients.showNotification(message,"warn",null,null,3000,true);
	}
}
