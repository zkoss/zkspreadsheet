/* NBooks.java

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

import org.zkoss.zss.ngmodel.impl.BookImpl;

/**
 * @author dennis
 * @since 3.5.0
 */
public class NBooks {

	public static NBook createBook(String bookName){
		NBook book = new BookImpl(bookName);
		return book;
	}
}
