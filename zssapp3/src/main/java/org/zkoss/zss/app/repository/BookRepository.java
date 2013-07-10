/* BookRepository.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/7/4 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.repository;

import java.io.IOException;
import java.util.List;

import org.zkoss.zss.api.model.Book;

/**
 * @author dennis
 *
 */
public interface BookRepository {

	List<BookInfo> list();
	
	Book load(BookInfo info) throws IOException;
	
	BookInfo save(BookInfo info,Book book) throws IOException;
	
	BookInfo saveAs(String name,Book book) throws IOException;
	
	boolean delete(BookInfo info) throws IOException;
}
