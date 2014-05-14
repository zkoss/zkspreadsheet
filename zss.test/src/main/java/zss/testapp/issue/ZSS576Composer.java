package zss.testapp.issue;

import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Window;

public class ZSS576Composer extends SelectorComposer<Window> {
	
	@Wire
	private Spreadsheet ss;
}