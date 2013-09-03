package zss.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.ui.event.HeaderEvent;
import org.zkoss.zul.Label;
import org.zkoss.zul.Vlayout;

@SuppressWarnings("serial")
public class WrapComposer extends SelectorComposer<Component>{

	@Wire
	private Vlayout msgArea;

	@Listen("onHeaderSize = spreadsheet")
	public void printHeaderResizeInfo(HeaderEvent event){
		if (event.getType() == HeaderEvent.LEFT_HEADER){
			Label messageLabel = new Label("resize row" + (event.getIndex()+1)+" to "+event.getData());
			msgArea.insertBefore(messageLabel, msgArea.getFirstChild());
		}
	}
	
	
}
