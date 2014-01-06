package org.zkoss.zss.ngmodel.impl;

import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/**
 * a collector to collect dependent when doing model manipulation
 * @author dennis
 *
 */
public class DependentCollector {

	static ThreadLocal<DependentCollector>  current = new ThreadLocal<DependentCollector>();
	
	private Set<Ref> dependents;
	
	public DependentCollector(){
	}
	public static DependentCollector setCurrent(DependentCollector ctx){
		DependentCollector old = current.get();
		current.set(ctx);
		return old;
	}
	
	public static DependentCollector getCurrent(){
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
