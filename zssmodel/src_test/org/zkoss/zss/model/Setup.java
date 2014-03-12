package org.zkoss.zss.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.Properties;

import org.zkoss.lang.Library;

public class Setup {
	
	static {
		Properties props = new Properties();
		String config = "zssmodel.test.properties";
		try {
			props.load(new FileInputStream(config));
		} catch (FileNotFoundException e) {
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		for(Iterator<?> it = props.keySet().iterator(); it.hasNext();) {
			String key = (String) it.next();
			Library.setProperty(key, props.getProperty(key));
		}
	}
	
	public static void touch() {
	};
}