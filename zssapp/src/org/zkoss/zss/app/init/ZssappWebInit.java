/* ZssappWebInit.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 21, 2011 12:46:51 PM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.init;

import org.zkoss.zk.ui.WebApp;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.metainfo.ComponentDefinition;
import org.zkoss.zk.ui.metainfo.LanguageDefinition;
import org.zkoss.zk.ui.util.WebAppInit;

/**
 * @author sam
 *
 */
public class ZssappWebInit implements WebAppInit {


	public void init(WebApp wapp) throws Exception {
		
		//use ZK PE's widget implementation if available
		LanguageDefinition langDef = LanguageDefinition.getByExtension("zul");
		ComponentDefinition colorButtonDef = langDef.getComponentDefinitionIfAny("colorbutton");
		if (colorButtonDef != null && WebApps.getFeature("pe")) {
			colorButtonDef.setDefaultWidgetClass("zssappex.Colorbutton");
		}
	}

}
