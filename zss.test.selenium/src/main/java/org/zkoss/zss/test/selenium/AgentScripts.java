package org.zkoss.zss.test.selenium;

import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;


public class AgentScripts {

	static volatile private AgentScripts instance;
	String staticJsContent;
//	
//	Map<String,String> jspool = new HashMap<String,String>();
	
	private AgentScripts(){
		try {
			init();
		} catch (Exception e) {
			if(e instanceof RuntimeException){
				throw (RuntimeException)e;
			}
			throw new RuntimeException(e.getMessage(),e);
		}
	}
	
	
	
	private void init() throws UnsupportedEncodingException, IOException {
		InputStream is = getClass().getResourceAsStream("/agent.js");
		if (is == null) {
			System.out.println("null");
		}
		try{
			staticJsContent = new String(Files.readAll(is),"UTF-8");
		}finally{
			if(is!=null){
				is.close();
			}
		}
	}
	
	public String getScript(){
		return staticJsContent;
	}

	public static AgentScripts instance(){
		if(instance==null){
			synchronized(AgentScripts.class){
				if(instance==null){
					instance = new AgentScripts();
				}
			}
		}
		return instance;
	}
	
//	public static void main(String[] args){
//		WidgetScriptPool.instance();
//	}
}
