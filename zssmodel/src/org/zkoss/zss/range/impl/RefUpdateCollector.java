package org.zkoss.zss.range.impl;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.model.sys.dependency.Ref;

/**
 * a collector to collect ref when doing model manipulation
 * @author dennis
 *
 */
public class RefUpdateCollector {

	static ThreadLocal<RefUpdateCollector>  current = new ThreadLocal<RefUpdateCollector>();
	
	private Set<Ref> refs;
	
	public RefUpdateCollector(){
	}
	public static RefUpdateCollector setCurrent(RefUpdateCollector ctx){
		RefUpdateCollector old = current.get();
		current.set(ctx);
		return old;
	}
	
	public static RefUpdateCollector getCurrent(){
		return current.get();
	}

	public void addRef(Ref ref) {
		if(refs==null){
			refs = new LinkedHashSet<Ref>();
		}
		refs.add(ref);
	}
	
	public Set<Ref> getRefs(){
		return refs==null?Collections.EMPTY_SET:Collections.unmodifiableSet(refs);
	}

	public void addRefs(Set<Ref> refs) {
		if(this.refs==null){
			this.refs = new LinkedHashSet<Ref>();
		}
		this.refs.addAll(refs);
	}
	
}
