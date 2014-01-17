package org.zkoss.zss.ngapi.impl;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/**
 * a collector to collect dependent when doing model manipulation
 * @author dennis
 *
 */
public class DependentUpdateCollector {

	static ThreadLocal<DependentUpdateCollector>  current = new ThreadLocal<DependentUpdateCollector>();
	
	private Set<Ref> dependents;
	
	public DependentUpdateCollector(){
	}
	public static DependentUpdateCollector setCurrent(DependentUpdateCollector ctx){
		DependentUpdateCollector old = current.get();
		current.set(ctx);
		return old;
	}
	
	public static DependentUpdateCollector getCurrent(){
		return current.get();
	}

	public void addDependent(Ref dependent) {
		if(dependents==null){
			dependents = new LinkedHashSet<Ref>();
		}
		dependents.add(dependent);
	}
	
	public Set<Ref> getDependents(){
		return dependents==null?Collections.EMPTY_SET:Collections.unmodifiableSet(dependents);
	}

	public void addDependents(Set<Ref> dependents) {
		if(this.dependents==null){
			this.dependents = new LinkedHashSet<Ref>();
		}
		this.dependents.addAll(dependents);
	}
	
}
