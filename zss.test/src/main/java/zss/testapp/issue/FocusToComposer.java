package zss.testapp.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Intbox;

public class FocusToComposer extends SelectorComposer<Component>{

	private static final long serialVersionUID = 1L;
	
	@Wire("spreadsheet")
	private Spreadsheet spreadsheet;
	@Wire("#row")
	private Intbox rowBox;
	@Wire("#col")
	private Intbox colBox;
	
	
	@Listen("onClick = button")
	public void focusTo(){
		spreadsheet.focusTo(rowBox.getValue(), colBox.getValue());
	}

}
