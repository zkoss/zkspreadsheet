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
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueue;
import org.zkoss.zk.ui.event.EventQueues;
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
	private final Map<String, RefSheet> _sheetRefs;
	private final int _maxrow;
	private final int _maxcol;
	private final ConcurrentMap<String, Ref> _variableRefs;
	private final EventQueue _queue;
	private final String _queueName;
	private static int _lastQueueId = -1;
	
	/**
	 * Internal use only.
	 */
	public RefBookImpl(String bookname, int maxrow, int maxcol) {
		_bookname = bookname;
		_sheetRefs = new HashMap<String, RefSheet>(3);
		_variableRefs = new ConcurrentHashMap<String, Ref>(4);
		_maxrow = maxrow;
		_maxcol = maxcol;
		_queueName = new StringBuilder().append("zss.").append(nextQueueId())
				.append(".").append(_bookname).toString();//_bookname is optional
		EventQueue tmp = null;
		try {
			tmp = EventQueues.lookup(getQueueName());
		} catch(IllegalStateException ex) {
			//ignore for zsstest case(No execution)
		}
		_queue = tmp;
	}
	
	private static synchronized int nextQueueId(){
		return ++_lastQueueId;
	}
	
	protected String getQueueName(){
		return _queueName;
	}

	
	@Override
	public RefSheet getOrCreateRefSheet(String sheetname) {
		RefSheet sheet = _sheetRefs.get(sheetname);
		if (sheet == null) {
			sheet = new RefSheetImpl(this, sheetname);
			_sheetRefs.put(sheetname, sheet);
		}
		return sheet;
	}

	@Override
	public RefSheet getRefSheet(String sheetname) {
		return _sheetRefs.get(sheetname);
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
		final EventQueue que = getEventQueue();
		if (que != null) {
			que.publish(event);
		}
	}
	
	protected EventQueue getEventQueue() {
		return _queue;
	}

	@Override
	public RefSheet removeRefSheet(String sheetname) {
		return _sheetRefs.remove(sheetname);
	}

	@Override
	public void setSheetName(String oldsheetname, String newsheetname) {
		final RefSheet sheet = _sheetRefs.remove(oldsheetname);
		if (sheet != null) {
			((RefSheetImpl)sheet).setSheetName(newsheetname);
			_sheetRefs.put(newsheetname, sheet);
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
		final Ref ref = new VarRefImpl(name, dummy);
		final Ref ret = _variableRefs.putIfAbsent(name, ref);
		return ret != null ? ret : ref;
	}

	@Override
	public Ref removeVariableRef(String name) {
		return _variableRefs.remove(name);
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

	protected String _scope;
	@Override
	public void setShareScope(String scope) {
		_scope = scope;
	}
	
	@Override
	public String getShareScope() {
		return _scope;
	}
}
