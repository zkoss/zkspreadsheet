package zss.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventQueue;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.ui.Spreadsheet;

public class Composer1168 extends SelectorComposer<Component> {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	@Wire
	Spreadsheet ss;
	
	private EventQueue<Event> queue = EventQueues.lookup("myQueue");
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		queue.subscribe(new EventListener<Event>() {
			
			@Override
			public void onEvent(Event e) throws Exception {
				if (e.getName().equals("onPopulate")){
					clickMerge();
					queue.publish(new Event("complete"));
				}
			}
		}, true);
		
		queue.subscribe(new EventListener<Event>() {
			@Override
			public void onEvent(Event e) throws Exception {
				if (e.getName().equals("complete")){
					Ranges.range(ss.getSelectedSheet()).notifyChange();
					Clients.clearBusy();
				}
			}
		});
		//publish(null);
	}
	
	@Listen("onClick= #eventQueue")
	public void publish(Event e){
		queue.publish(new Event("onPopulate"));
		Clients.showBusy("populating");
	}
	
	@Listen("onClick= #doRefresh")
	public void merge(Event e){
		Range range = Ranges.range(ss.getSelectedSheet(), 0, 0, 10, 9);
		range.setAutoRefresh(false);
		range.merge(false);
		Ranges.range(ss.getSelectedSheet()).notifyChange();
	}
	
	public void clickMerge(){
		Range range = Ranges.range(ss.getSelectedSheet(), 0, 0, 10, 9);
		CellOperationUtil.merge(range, false); //or range.merge(true);
//		Ranges.range(ss.getSelectedSheet()).notifyChange();
	}
}
