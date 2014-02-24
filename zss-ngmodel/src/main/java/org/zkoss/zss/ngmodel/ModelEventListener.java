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
package org.zkoss.zss.ngmodel;

import java.io.Serializable;


/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface ModelEventListener extends Serializable{

	public void onEvent(ModelEvent event);
}
