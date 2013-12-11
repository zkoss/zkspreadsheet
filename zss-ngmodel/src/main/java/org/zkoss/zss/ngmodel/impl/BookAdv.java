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

import java.io.Serializable;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class BookAdv implements NBook,Serializable{
	private static final long serialVersionUID = 1L;
	
	/**
	 * Optimize CellStyle, usually called when export book. 
	 * @return
	 */
	/*package*/ abstract List<NCell> optimizeCellStyle();
	
	/*package*/ abstract String nextObjId(String type);
	
	/*package*/ abstract void onModelInternalEvent(ModelInternalEvent event);
	
	/*package*/ abstract void sendModelInternalEvent(ModelInternalEvent event);
	
	public abstract void sendModelEvent(ModelEvent event);
	
	
	ModelInternalEvent createModelInternalEvent(String name, Object... data){
		return ModelInternalEvents.createModelInternalEvent(name,this,data);
	}

}
