/* ExportComposer.java

	Purpose:
		
	Description:
		
	History:
		November 05, 5:53:16 PM     2010, Created by Ashish Dasnurkar

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
 */
package zss.issue;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueue;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author hawk
 * 
 */
public class SetMassiveCellsComposer1115 extends SelectorComposer {
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
					populateData();
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

	}
	
	@Listen("onClick= #populate")
	public void start(Event e){
		Events.echoEvent("onPopulate", ss, null);
		Clients.showBusy("populating");
	}
	
	@Listen("onPopulate= #ss")
	public void populate(Event e) throws InterruptedException{
		populateData();
		Clients.clearBusy();
	}

	private void populateData() {
		Sheet sheet = ss.getSelectedSheet();
		for (int column  = 0 ; column < 20 ; column++){
			for (int row = 0 ; row <800 ; row++ ){
				Range range = Ranges.range(sheet, row, column);
				range.setAutoRefresh(false);
				range.getCellData().setEditText(row+", "+column);
				CellOperationUtil.applyBackgroundColor(range, "#FF00FF");
				CellOperationUtil.applyFontHeightPoints(range, 14);
				CellOperationUtil.applyAlignment(range, Alignment.CENTER);
			}
		}
		Ranges.range(ss.getSelectedSheet()).notifyChange();
		
	}
	
	@Listen("onClick= #publish")
	public void pushlish(Event e){
		queue.publish(new Event("onPopulate"));
		Clients.showBusy("populating");
	}
}
