package org.zkoss.zss.essential.util;

import org.zkoss.zk.ui.util.Clients;

public class ClientUtil {

	
	public static void showInfo(String msg){
		Clients.showNotification(msg,"info",null,null,3000,false);
	}
	public static void showWarn(String msg){
		Clients.showNotification(msg,"warn",null,null,3000,false);
	}
	public static void showError(String msg){
		Clients.showNotification(msg,"error",null,null,-1,true);
	}
}
