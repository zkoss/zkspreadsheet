package zss.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.ui.event.HeaderUpdateEvent;
import org.zkoss.zss.ui.event.HedaerType;
import org.zkoss.zul.Label;
import org.zkoss.zul.Vlayout;

@SuppressWarnings("serial")
public class ZSS_528Composer extends SelectorComposer<Component>{

	@Wire
	private Vlayout msgArea;

	@Listen("onHeaderUpdate = spreadsheet")
	public void printHeaderResizeInfo(HeaderUpdateEvent event){
		if (event.getType() == HedaerType.ROW){
			Label messageLabel = new Label("resize row" + (event.getIndex()+1)+" to "+event.getData());
			msgArea.insertBefore(messageLabel, msgArea.getFirstChild());
		}
	}
	
	
}
