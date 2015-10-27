/* SColorScale.java

	Purpose:
		
	Description:
		
	History:
		Oct 23, 2015 4:25:03 PM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.List;

/**
 * @author henri
 * @since 3.8.2
 */
public interface SColorScale {
	/** Returns the value object */
	List<SCFValueObject> getCFValueObjects();
	
	/** Returns the gradient colors */
	List<SColor> getColors();
}
