package zss.testapp.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.SheetOperationUtil;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Intbox;

public class FreezeComposer206 extends SelectorComposer<Component>{

	private static final long serialVersionUID = 1L;
	
	@Wire("spreadsheet")
	private Spreadsheet spreadsheet;
	@Wire("#row")
	private Intbox rowBox;
	
	
	@Listen("onClick = button")
	public void freeze(){
		spreadsheet.setRowfreeze(rowBox.getValue());
	}

}
