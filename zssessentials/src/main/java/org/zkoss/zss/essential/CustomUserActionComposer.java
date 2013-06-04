package org.zkoss.zss.essential;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.ui.Spreadsheet;

public class CustomUserActionComposer extends SelectorComposer<Component>{
	private static final long serialVersionUID = 1L;
	
	@Wire
	Spreadsheet ss;
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
	}
}
