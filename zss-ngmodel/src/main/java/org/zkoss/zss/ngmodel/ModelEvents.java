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
	public static final String ON_CHART_CONTENT_CHANGE = "onChartContentChange";
	public static final String ON_DATA_VALIDATION_CONTENT_CHANGE = "onDataValidationContentChange";
	
	public static final String ON_ROW_COLUMN_SIZE_CHANGE = "onRowColumnSizeChange";
	public static final String ON_AUTOFILTER_CHANGE = "onAutoFilterChange";
	public static final String ON_FREEZE_CHANGE = "onFreezeChange";
	
	public static final String PARAM_BOOK = "book";
	public static final String PARAM_SHEET = "sheet";
	public static final String PARAM_REGION = "region";
	public static final String PARAM_OBJECT_ID = "objid";
	
	
	
	public static ModelEvent createModelEvent(String name, NBook book){
		return createModelEvent0(name,book,null,null,null);
	}
	public static ModelEvent createModelEvent(String name, NBook book,Map data){
		return createModelEvent0(name,book,null,null,data);
	}
	public static ModelEvent createModelEvent(String name, NSheet sheet){
		return createModelEvent0(name,sheet.getBook(),sheet,null,null);
	}
	public static ModelEvent createModelEvent(String name, NSheet sheet,Map data){
		return createModelEvent0(name,sheet.getBook(),sheet,null,data);
	}
	public static ModelEvent createModelEvent(String name, NSheet sheet,CellRegion region){
		return createModelEvent0(name,sheet.getBook(),sheet,region,null);
	}
	public static ModelEvent createModelEvent(String name, NSheet sheet,CellRegion region,Map data){
		return createModelEvent0(name,sheet.getBook(),sheet,region,data);
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
	
	private static ModelEvent createModelEvent0(String name, NBook book, NSheet sheet,CellRegion region,Map data){
		Map<String,Object> datamap = new HashMap<String,Object>();
		if(book!=null){
			datamap.put(ModelEvents.PARAM_BOOK, book);
		}
		if(sheet!=null){
			datamap.put(ModelEvents.PARAM_SHEET, sheet);
		}
		if(region!=null){
			datamap.put(ModelEvents.PARAM_REGION, region);
		}
		if(data!=null){
			datamap.putAll(data);
		}
		ModelEvent event = new ModelEvent(name, datamap);
		return event;
	}
}
