package org.zkoss.zss.app.ui;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.repository.BookRepositoryFactory;

public class MainCtrl extends SelectorComposer<Component>{

	
	@Listen("onClick=#test")
	public void test(){
		BookRepository repository = BookRepositoryFactory.getInstance().getRepository();
		
		System.out.println(">>>>"+repository.list());
	}
}
