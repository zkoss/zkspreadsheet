package org.zkoss.zss.ngmodel.impl.sys;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import org.zkoss.zss.ngmodel.sys.dependency.DependencyTableAdv;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/* DependencyTableImpl.java

 Purpose:

 Description:

 History:
 Nov 22, 2013 Created by Pao Wang

 Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */

/**
 * @author Pao
 */
public class DependencyTableImpl extends DependencyTableAdv {

	/** Map<dependant, precedent> */
	private Map<Ref, Set<Ref>> map = new HashMap<Ref, Set<Ref>>();

	public DependencyTableImpl() {
	}

	@Override
	public void add(Ref dependant, Ref precedent) {
		Set<Ref> precedents = map.get(dependant);
		if(precedents == null) {
			precedents = new HashSet<Ref>();
			map.put(dependant, precedents);
		}
		precedents.add(precedent);
	}

	public void clear() {
		map.clear();
	}

	@Override
	public void clearDependents(Ref dependant) {
	}

	@Override
	public Set<Ref> getDependents(Ref precedent) {
		Set<Ref> result = new HashSet<Ref>();
		for(Entry<Ref, Set<Ref>> entry : map.entrySet()) {
			Ref target = entry.getKey();
			for(Ref pre : entry.getValue()) {
				if(isInside(pre, precedent)) {
					result.add(target);
				}
			}
		}
		return result;
	}

	/**
	 * @return true if b is inside a.
	 */
	private boolean isInside(Ref a, Ref b) {
		// TODO
		return false;
	}

	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		for(Entry<Ref, Set<Ref>> entry : map.entrySet()) {
			Ref target = entry.getKey();
			sb.append(target).append('\n');
			for(Ref pre : entry.getValue()) {
				sb.append('\t').append(pre).append('\n');
			}
		}
		return sb.toString();
	}

	@Override
	public void merge(DependencyTableAdv dependencyTable) {
		// TODO Auto-generated method stub
	}
}
