package org.zkoss.zss.app.ui;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueue;
import org.zkoss.zk.ui.event.EventQueues;
import org.zkoss.zk.ui.select.SelectorComposer;

public class CtrlBase<T extends Component> extends SelectorComposer<T>{

	private static final long serialVersionUID = 1L;
	private static final String QUEUE_NAME = "$ZSSAPP$";
	
	EventListener<DesktopEvent> desktopEventDispatcher = new EventListener<DesktopEvent>(){
		public void onEvent(DesktopEvent event) throws Exception {
			if(event.ignoreSelf && event.from==CtrlBase.this){
				return;
			}
			onDesktopEvent(event.getName(),event.getData());
		}
	};
	
	static class DesktopEvent extends Event{
		private static final long serialVersionUID = 1L;
		Object from;
		boolean ignoreSelf;
		
		public DesktopEvent(String name, Object from, boolean ignoreSelf, Object data) {
			super(name, null, data);
			this.from = from;
			this.ignoreSelf = ignoreSelf;
		}
	}
	
	public void doAfterCompose(T comp) throws Exception {
		super.doAfterCompose(comp);	
		EventQueue eq = EventQueues.lookup(QUEUE_NAME, "desktop", true);
		eq.subscribe(desktopEventDispatcher);
	}
	
	
	
	static protected class Entry {
		final String name;
		final Object value;
		public Entry(String name,Object value){
			this.name = name;
			this.value = value;
		}
	}
	
	static protected Entry newEntry(String name,Object value){
		Entry arg = new Entry(name,value);
		return arg;
	}
	
	static protected Map<String,Object> newMap(Entry... args){
		Map<String,Object> argm = new HashMap<String,Object>();
		for(Entry arg:args){
			argm.put(arg.name,arg.value);
		}
		return argm;
	}
	
	protected void pushDesktopEvent(String event,boolean ignoreSelf,Object data){
		EventQueue eq = EventQueues.lookup(QUEUE_NAME, "desktop", true);
		eq.publish(new DesktopEvent(event,this,ignoreSelf, data));
	}
	
	//should override this
	protected void onDesktopEvent(String event,Object data){
		//default do nothing.
	}
	
	
	protected void destroy(){
		EventQueue eq = EventQueues.lookup(QUEUE_NAME, "desktop", false);
		if(eq!=null){
			eq.unsubscribe(desktopEventDispatcher);
		}
	}
}
