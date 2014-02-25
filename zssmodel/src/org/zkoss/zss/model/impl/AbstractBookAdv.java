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
package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.Map;

import org.zkoss.zss.model.ModelEvent;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.sys.formula.EvaluationContributorContainer;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class AbstractBookAdv implements SBook,EvaluationContributorContainer,Serializable{
	private static final long serialVersionUID = 1L;
	
	public abstract void sendModelEvent(ModelEvent event);
	
	/*package*/ abstract String nextObjId(String type);
	
//	/*package*/ abstract void onModelInternalEvent(ModelInternalEvent event);
//	
//	/*package*/ abstract void sendModelInternalEvent(ModelInternalEvent event);

	/*package*/ abstract void setBookSeries(SBookSeries bookSeries);

}
