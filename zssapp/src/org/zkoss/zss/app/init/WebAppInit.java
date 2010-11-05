/* WebAppInit.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Oct 28, 2010 9:38:09 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.init;

import org.zkoss.zk.ui.WebApp;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.metainfo.ComponentDefinition;
import org.zkoss.zk.ui.metainfo.LanguageDefinition;

/**
 * @author Sam
 *
 */
public class WebAppInit implements org.zkoss.zk.ui.util.WebAppInit{


	public void init(WebApp webapp) throws Exception {
		if (WebApps.getFeature("pe")) {
			LanguageDefinition langDef = LanguageDefinition.getByExtension("zul");
			
			ComponentDefinition colorBtnDef = langDef.getComponentDefinition("colorbutton");
			if (colorBtnDef != null) {
				colorBtnDef.setImplementationClass("org.zkoss.zss.app.zul.Colorbutton");
				colorBtnDef.setDefaultWidgetClass("zssapp.Colorbutton");
			}
		}
	}
}