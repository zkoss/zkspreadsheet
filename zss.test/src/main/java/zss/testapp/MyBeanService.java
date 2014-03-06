package zss.testapp;

import org.zkoss.zk.ui.Executions;

public class MyBeanService {
	static private MyBeanService myBeanService = null;

	private MyBeanService(){
		
	}
	
	public Object get(String name){
		if ("myBean".equals(name) && Executions.getCurrent().getSession().getAttribute(name)==null){
			//create beans if no exist
			Executions.getCurrent().getSession().setAttribute("myBean",new AssetsBean());
		}
		return Executions.getCurrent().getSession().getAttribute(name);
	}
	
	
	
	static public MyBeanService getMyBeanService(){
		if (myBeanService == null){
			myBeanService = new MyBeanService();
		}
		
		return myBeanService;
	}
}
