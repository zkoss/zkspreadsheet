/* NBookSeries.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/11/14 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel;

import java.util.Set;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.lang.Strings;

/**
 * @author dennis
 * @since 3.5.0
 */
public abstract class NBookSeriesBuilder {
	
	private static NBookSeriesBuilder _instance;
	
	public static NBookSeriesBuilder getInstance(){
		
		if(_instance==null){
			synchronized(NBookSeriesBuilder.class){
				if(_instance==null){
					String clz = Library.getProperty("org.zkoss.zss.api.BookSeriesBuilder.class");
					if (!Strings.isEmpty(clz)) {
						try {
							_instance = (NBookSeriesBuilder) Classes.forNameByThread(clz).newInstance();
						} catch (Exception e) {
							throw new RuntimeException(e.getMessage(),e);
						}
					}else{
						_instance = new NBookSeriesBuilder() {
							@Override
							public void buildBookSeries(Set<NBook> books) {
								throw new RuntimeException("not implemented");
							}
							@Override
							public void buildBookSeries(NBook[] books) {
								throw new RuntimeException("not implemented");
							}
						};
					}
				}
			}
		}
		return _instance;
	}
	
	
	abstract public void buildBookSeries(Set<NBook> books);
	abstract public void buildBookSeries(NBook... books);
}
