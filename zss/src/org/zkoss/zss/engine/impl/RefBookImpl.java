/* RefBookImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 8, 2010 11:53:00 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.engine.impl;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueue;
import org.zkoss.zk.ui.event.impl.DesktopEventQueue;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.RefSheet;

/**
 * Implementation of {@link RefBook} that manage the {@link RefSheet}s associated in this RefBook.
 * @author henrichen
 *
 */
public class RefBookImpl implements RefBook {
	private final String _bookname;
	private final Map<String, RefSheet> _sheets;
	private final int _maxrow;
	private final int _maxcol;
	private final Map<String, Ref> _variableRefs;
	
	/**
	 * Internal use only.
	 */
	public RefBookImpl(String bookname, int maxrow, int maxcol) {
		_bookname = bookname;
		_sheets = new HashMap<String, RefSheet>(3);
		_variableRefs = new HashMap<String, Ref>(4);
		_maxrow = maxrow;
		_maxcol = maxcol;
	}
	
	@Override
	public RefSheet getOrCreateRefSheet(String sheetname) {
		RefSheet sheet = _sheets.get(sheetname);
		if (sheet == null) {
			sheet = new RefSheetImpl(this, sheetname);
			_sheets.put(sheetname, sheet);
		}
		return sheet;
	}

	@Override
	public RefSheet getRefSheet(String sheetname) {
		return _sheets.get(sheetname);
	}
	
	@Override
	public String getBookName() {
		return _bookname;
	}
	
	@Override
	public void subscribe(EventListener listener) {
		getEventQueue().subscribe(listener);
	}
	
	@Override
	public void unsubscribe(EventListener listener) {
		getEventQueue().unsubscribe(listener);
	}
	
	@Override
	public void publish(Event event) {
		getEventQueue().publish(event);
	}
	
	//@see http://en.wikipedia.org/wiki/Double-checked_locking
	private volatile EventQueue _queue; 
	private EventQueue getEventQueue() {
		EventQueue queue = _queue;
		if (queue == null) {
            synchronized(this) {
            	queue = _queue; 
    			if (queue == null) {
					final String clsnm = Library.getProperty("org.zkoss.zss.engine.EventQueue.class");
					if (clsnm != null) {
						try {
							final Object o = Classes.newInstanceByThread(clsnm);
							//try zkex first
							if (!(o instanceof EventQueue))
								throw new UiException(o.getClass().getName()+" must implement "+EventQueue.class.getName());
							queue = (EventQueue)o;
						} catch (UiException ex) {
							throw ex;
						} catch (Throwable ex) {
							throw UiException.Aide.wrap(ex, "Unable to load "+clsnm);
						}
						
					}
					if (queue == null)
						queue = new DesktopEventQueue(); 
    				_queue = queue;
    			}
            }
		}
		return queue;
	}

	@Override
	public RefSheet removeRefSheet(String sheetname) {
		return _sheets.remove(sheetname);
	}

	@Override
	public void setSheetName(String oldsheetname, String newsheetname) {
		final RefSheet sheet = _sheets.remove(oldsheetname);
		if (sheet != null) {
			_sheets.put(newsheetname, sheet);
		}
	}
	@Override
	public int getMaxrow() {
		return _maxrow;
	}
	
	@Override
	public int getMaxcol() {
		return _maxcol;
	}

	@Override
	public Ref getOrCreateVariableRef(String name, RefSheet dummy) {
		Ref ref = _variableRefs.get(name);
		if (ref == null) {
			synchronized(this) {
				if (ref == null) {
					ref = new VarRefImpl(name, dummy);
					_variableRefs.put(name, ref);
				}
			}
		}
		return ref;
	}

	@Override
	public Ref removeVariableRef(String name) {
		synchronized(this) {
			return _variableRefs.remove(name);
		}
	}
	
	@SuppressWarnings("unchecked")
	@Override
	public Set<Ref>[] getBothDependents(String name) {
		final Ref ref = _variableRefs.get(name);
		if (ref != null) {
			final Set<Ref> last = new HashSet<Ref>();
			final Set<Ref> all = new HashSet<Ref>();
			DependencyTrackerHelper.getBothDependents(ref.getDependents(), all, last);
			return (Set<Ref>[]) new Set[] {last, all};
		}
		return null;
	}

}
