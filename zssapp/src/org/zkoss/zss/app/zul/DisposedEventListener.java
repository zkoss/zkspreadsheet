/* DisposedEventListener.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 11:54:45 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.zk.ui.event.EventListener;

/**
 * @author Ian Tsai
 *
 */
public interface DisposedEventListener extends EventListener {
	
	public boolean isDisposed();

}
