/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.sys.dependency;

import java.util.List;
import java.util.Set;

/**
 * 
 * NodeA --- depends on ---> NodeB.
 * A is B's dependent, B is A's precedent.
 * When B changes , should call {@link #getDependents(Ref)} of B to create notification of A
 * When A been clear or deleted, should call {@link #clearDependents(Ref)} of A to clear tracking data
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface DependencyTable {

	public Set<Ref> getDependents(Ref precedent);
	
	public void clearDependents(Ref dependant);
}
