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

import java.util.HashMap;
import java.util.Map;


/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ModelEvents {
	
	public static final String ON_CELL_CONTENT_CHANGE = "onCellChange";
	
	public static final String PARAM_BOOK = "book";
	public static final String PARAM_SHEET = "sheet";
	public static final String PARAM_REGION = "region";
	
	
	public static ModelEvent createModelEvent(String name, NBook book,Object... data){
		return createModelEvent(name,book,null,null,data);
	}
	public static ModelEvent createModelEvent(String name, NSheet sheet, Object... data){
		return createModelEvent(name,sheet.getBook(),sheet,null,data);
	}
	public static ModelEvent createModelEvent(String name, NBook book, NSheet sheet,CellRegion region,Object... data){
		Map<String,Object> datamap = new HashMap<String,Object>();
		if(book!=null){
			datamap.put(ModelEvents.PARAM_BOOK, book);
		}
		if(sheet!=null){
			datamap.put(ModelEvents.PARAM_SHEET, sheet);
		}
		if(sheet!=null){
			datamap.put(ModelEvents.PARAM_REGION, region);
		}
		if(datamap!=null){
			if(data.length%2 != 0){
				throw new IllegalArgumentException("event data must be key,value pair");
			}
			for(int i=0;i<data.length;i+=2){
				if(!(data[i] instanceof String)){
					throw new IllegalArgumentException("event data key must be string");
				}
				datamap.put((String)data[i],data[i+1]);
			}
		}
		ModelEvent event = new ModelEvent(name, datamap);
		return event;
	}
}
