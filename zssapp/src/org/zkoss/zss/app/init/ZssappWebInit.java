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

import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Locale;

import org.zkoss.lang.Strings;
import org.zkoss.util.resource.Labels;
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

		// use ZK PE's widget implementation if available
		LanguageDefinition langDef = LanguageDefinition.getByExtension("zul");
		ComponentDefinition colorButtonDef = langDef
				.getComponentDefinitionIfAny("colorbutton");
		if (colorButtonDef != null && WebApps.getFeature("pe")) {
			colorButtonDef.setDefaultWidgetClass("zssappex.Colorbutton");
		}

		initResources(wapp);
	}

	private void initResources(WebApp wapp) throws URISyntaxException,
			MalformedURLException {
		ClassLoader locator = Thread.currentThread().getContextClassLoader();

		ArrayList<URL> urls = new ArrayList<URL>();
		// FIXME: remove these hard code here
		URL i3labelURL = locator.getResource("web/zssapp/labels/i3-label.properties");
		if (i3labelURL != null)
			urls.add(i3labelURL);
		URL tw = locator.getResource("web/zssapp/labels/i3-label_zh_TW.properties");
		if (tw != null)
			urls.add(tw);
		urls.add(i3labelURL);

		Labels.register(new ZssappLabelLocator(urls));
	}

	 private class ZssappLabelLocator implements org.zkoss.util.resource.LabelLocator {
		 private URL _default; 
		 private HashMap<String, URL> urlMap = new HashMap<String, URL>();
		 
		 ZssappLabelLocator(ArrayList<URL> urls) {
			 mapLocaleToURL(urls);
		 }
		 @Override
		 public URL locate(Locale locale) throws Exception {
			 return bestMatchURL(locale);
		 }
		 
		 private URL bestMatchURL(Locale locale) {
			 if (locale == null) {
				 return _default;
			 } else {
				 URL url = (URL) urlMap.get(locale.toString());
				 if (url != null) {
					 return url;
				 } else {
					 for (Iterator<String> i = urlMap.keySet().iterator(); i.hasNext() ; ) {
						 String l = i.next();
						 if (l.startsWith(locale.toString()))
							 return urlMap.get(l);
					 }
					 return _default;
				 }
			 }
		 }
		 
		 private void mapLocaleToURL(ArrayList<URL> urls) {
			 int startIdx = 0;
			 int endIdx = 0;
			 for (URL url : urls) {
				 final String s = url.toString();
				 endIdx = s.lastIndexOf(".properties");
				 startIdx = s.lastIndexOf("i3-label", endIdx) + "i3-label".length();
				 String test = s.substring(startIdx, endIdx);
				 if (Strings.isEmpty(test)) {
					 _default = url;
				 } else {
					 String locale = fixLocale(s.substring(startIdx, endIdx));
					 urlMap.put(locale, url);
				 }
			 }
		 }
		 private String fixLocale(String locale) {
			 return locale.startsWith("_") ? locale.substring(1) : locale;
		 }
	 }
}
