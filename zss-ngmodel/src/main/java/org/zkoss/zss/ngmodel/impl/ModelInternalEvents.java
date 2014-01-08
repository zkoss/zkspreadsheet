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
package org.zkoss.zss.ngmodel.impl;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NSheet;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ModelInternalEvents {

	public static final String ON_SHEET_ADDED = "onSheetAdded";
	public static final String ON_SHEET_RENAMED = "onSheetRenamed";
	public static final String ON_SHEET_DELETED = "onSheetDeleted";
	public static final String ON_SHEET_MOVED = "onSheetMoved";
	
	public static final String ON_ROW_INSERTED = "onRowInserted";
	public static final String ON_ROW_DELETED = "onRowDeleted";
	public static final String ON_COLUMN_INSERTED = "onColumnInserted";
	public static final String ON_COLUMN_DELETED = "onColumnDeleted";
	

	public static final String PARAM_SHEET_OLD_NAME = "sheetName";
	public static final String PARAM_SHEET_OLD_INDEX = "sheetIdx";
	public static final String PARAM_ROW_INDEX = "rowIdx";
	public static final String PARAM_COLUMN_INDEX = "columnIdx";
	public static final String PARAM_SIZE = "size";
	
	public static ModelInternalEvent createModelInternalEvent(String name, NBook book){
		return createModelInternalEvent(name,book,null,null);
	}
	public static ModelInternalEvent createModelInternalEvent(String name, NBook book,Map data){
		return createModelInternalEvent(name,book,null,data);
	}
	public static ModelInternalEvent createModelInternalEvent(String name, NSheet sheet){
		return createModelInternalEvent(name,sheet.getBook(),sheet,null);
	}
	public static ModelInternalEvent createModelInternalEvent(String name, NSheet sheet,Map data){
		return createModelInternalEvent(name,sheet.getBook(),sheet,data);
	}
	public static ModelInternalEvent createModelInternalEvent(String name, NBook book, NSheet sheet){
		return createModelInternalEvent(name,book,sheet,null);
	}
	public static ModelInternalEvent createModelInternalEvent(String name, NBook book, NSheet sheet,Map data){
		Map<String,Object> datamap = new HashMap<String,Object>();
		if(book!=null){
			datamap.put(ModelEvents.PARAM_BOOK, book);
		}
		if(sheet!=null){
			datamap.put(ModelEvents.PARAM_SHEET, sheet);
		}
		if(data!=null){
			datamap.putAll(data);
		}
		ModelInternalEvent event = new ModelInternalEvent(name, datamap);
		return event;
	}
	
	public static Map createDataMap(Object... data){
		if(data!=null){
			if(data.length%2 != 0){
				throw new IllegalArgumentException("event data must be key,value pair");
			}
			Map<String,Object> datamap = new HashMap<String,Object>();
			for(int i=0;i<data.length;i+=2){
				if(!(data[i] instanceof String)){
					throw new IllegalArgumentException("event data key must be string");
				}
				datamap.put((String)data[i],data[i+1]);
			}
			return datamap;
		}
		return null;
	}
}
