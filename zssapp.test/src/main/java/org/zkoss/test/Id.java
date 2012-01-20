/* Id.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 9:35:03 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import com.google.common.base.Preconditions;

/**
 * @author sam
 *
 */
public class Id {
	
	private String id;
	public Id (String id){
		this.id = Preconditions.checkNotNull(id);
	}
	
	@Override
	public String toString() {
		return id;
	}
}
