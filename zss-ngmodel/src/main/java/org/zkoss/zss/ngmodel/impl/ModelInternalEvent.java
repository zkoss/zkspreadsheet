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
package org.zkoss.zss.ngmodel.impl;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
/**
 * Model internal events
 * @author dennis
 * @since 3.5.0
 */
/*package*/ class ModelInternalEvent extends ModelEvent{

	private String name;
	
	private Map<String,Object> data;
	
	public ModelInternalEvent(String name){
		super(name);
	}
	
	public ModelInternalEvent(String name,Map<String,Object> data){
		super(name,data);
	}
}
