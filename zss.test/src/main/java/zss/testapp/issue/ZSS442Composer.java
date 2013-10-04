/* ZSS442Composer.java

	Purpose:
		
	Description:
		
	History:
		Oct 4, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package zss.testapp.issue;

import java.util.Arrays;
import java.util.LinkedList;
import java.util.Queue;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * A workaround for ZSS-442.
 * 
 * When we invoke multiple range API in single AU, the process order is:
 *   1. process every range API call and post multiple asynchronous event
 *       NOTE that these calls modify model immediately
 *   2. process events from invoking range API
 *       these listeners notify client according to these event
 *       NOTE that model has already became to the final status
 *   3. client receive these responses and update cells, but cell data is based on final status.
 *       every update operations MUST be based on the status at that time not final status.
 *       this causes conflict and make client is inconsistent with server side.
 *       
 * This workaround separates every range API call into individual actions.
 * The action queue will perform these actions one by one. 
 * It can invoke a range API after previous range API call has processed and already posted corresponding ZK events.
 * @author Pao
 */
public class ZSS442Composer extends SelectorComposer<Component> {
	private static final long serialVersionUID = 1L;

	@Wire
	private Spreadsheet ss;

	@Listen("onClick = #workaround")
	public void doProcess() {
		new ActionQueue().add(new Action() {
			public void perform() {
				Ranges.range(ss.getSelectedSheet(), "G3").delete(DeleteShift.LEFT);
			}
		}).add(new Action() {
			public void perform() {
				Ranges.range(ss.getSelectedSheet(), "F3").insert(InsertShift.RIGHT,
						InsertCopyOrigin.FORMAT_RIGHT_BELOW);
			}
		}).add(new Action() {
			public void perform() {
				Ranges.range(ss.getSelectedSheet(), "F").toColumnRange().delete(DeleteShift.LEFT);
			}
		}).submit();
	}
}

interface Action {
	void perform();
}

class ActionQueue implements EventListener<Event> {
	private final static String ID = "zss.range.action";
	private Queue<Action> queue;

	public ActionQueue() {
		queue = new LinkedList<Action>();
		EventQueues.lookup(ID).subscribe(this);
	}

	public ActionQueue add(Action action) {
		queue.add(action);
		return this;
	}

	public ActionQueue Add(Action[] actions) {
		queue.addAll(Arrays.asList(actions));
		return this;
	}

	public void submit() {
		EventQueues.lookup(ID).publish(new Event(ID));
	}

	@Override
	public void onEvent(Event event) throws Exception {
		if(!ID.equals(event.getName())) { // just in case
			return;
		}
		try {
			if(!queue.isEmpty()) {
				queue.remove().perform();
				EventQueues.lookup(ID).publish(new Event(ID));
			}
		} catch(Exception e) {
			queue.clear();
			throw e;
		}
	}
}
