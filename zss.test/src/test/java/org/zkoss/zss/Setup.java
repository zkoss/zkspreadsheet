package org.zkoss.zss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.Properties;

import org.zkoss.lang.Library;

public class Setup {
	
	static {
		Properties props = new Properties();
		try {
			props.load(new FileInputStream("zss.test.properties"));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		for(Iterator<?> it = props.keySet().iterator(); it.hasNext();) {
			String key = (String) it.next();
			Library.setProperty(key, props.getProperty(key));
		}
	}
	
	public static void touch() {};
	
	static private File temp; 
	
	/**
	 * get the temp file, it is always the same one in same testing vm and will be deleted after testing.
	 * @return
	 */
	public static synchronized File getTempFile(){
		if(temp!=null){
			return temp;
		}
		temp = getTempFile("zsstest","");
		temp.deleteOnExit();
		return temp;
	}
	/**
	 * get a temp file, with a prefix aod postfix, the provided file will not be deleted after test.
	 */
	public static synchronized File getTempFile(String prefix,String postfix){
		File tempFolder = new File(System.getProperty("java.io.tmpdir"));
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmssSSS");
		File file = null;
		do{
			file = new File(tempFolder,prefix+"-"+sdf.format(new java.util.Date())+postfix);
			try {
				Thread.sleep(10);
			} catch (InterruptedException e) {}
		}while(file.exists());
		return file;
	}
}
