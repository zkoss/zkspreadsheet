package org.zkoss.zss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.Locale;
import java.util.Properties;
import java.util.Stack;

import org.zkoss.lang.Library;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.util.Locales;

public class Setup {
	
	static {
		Properties props = new Properties();
		try {
			props.load(new FileInputStream("zss.test.properties"));
		} catch (FileNotFoundException e) {
			System.out.println("** Have to copy file \"zsscml/cmldoc/zss.test.properties\" to folder \"" + System.getProperty("user.dir") + "\" first");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		for(Iterator<?> it = props.keySet().iterator(); it.hasNext();) {
			String key = (String) it.next();
			Library.setProperty(key, props.getProperty(key));
		}
	}
	
	public static void touch() {
		if(ZssContext.getCurrent()==null){
			Locale def = Locale.TAIWAN; //TODO read local from config
			ZssContext.setThreadLocal(new ZssContext(def,-1));
		}
		
	};
	
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
	
	//in 3.5 , we count on zk Locales
	static ThreadLocal<Stack<Locale>> zssLocale = new ThreadLocal<Stack<Locale>>();
	static{
		zssLocale.set(new Stack<Locale>());
	}
	
	public static void pushZssLocale(Locale l){
		Locale old = Locales.getCurrent();
		if(old!=null){
			zssLocale.get().push(old);
		}
		Locales.setThreadLocal(l);
	}
	
	public static void popZssLocale(){
		if(zssLocale.get().isEmpty()){
			System.out.print("WARN: No zss context to pop, please check your call stack");
			return;
		}
		Locale old = zssLocale.get().pop();
		Locales.setThreadLocal(old);
	}
	
	public static Locale getZssLocale(){
		if(zssLocale.get().isEmpty()){
			return null;
		}
		return Locales.getCurrent();
	}
	
	//set current content locale
	public static void setZssLocale(Locale l){
		if(!zssLocale.get().isEmpty()){
			popZssLocale();
		}
		pushZssLocale(l);
	}
}
