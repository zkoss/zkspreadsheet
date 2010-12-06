/* Preconditions.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 4:24:40 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */
package org.zkoss.zss.app.base;

import org.zkoss.zk.ui.UiException;

/**
 * @author Sam
 * 
 */
public final class Preconditions {
	private Preconditions() {
	}

	/**
	 * Ensures that an object reference passed as a parameter to the calling
	 * method is not null.
	 * 
	 * @param reference an object reference
	 * @return the non-null reference that was validated
	 * @throws UiException if {@code reference} is null
	 */
	public static <T> T checkNotNull(T reference, Object errorMessage) {
		if (reference == null) {
			throw new UiException(String.valueOf(errorMessage));
		}
		return reference;
	}
}
