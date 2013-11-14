/* NBook.java

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
package org.zkoss.zss.ngmodel.impl;

import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NSheet;

/**
 * @author dennis
 *
 */
public class BookImpl implements NBook{
	
	protected List<SheetImpl> sheets = new LinkedList<SheetImpl>();
	
	
	public BookImpl(){
		createSheet(null);
	}
	
	public NSheet getSheetAt(int i){
		return sheets.get(i);
	}
	
	public int getNumOfSheet(){
		return sheets.size();
	}
	
	public NSheet getSheetByName(String name){
		for(NSheet sheet:sheets){
			if(sheet.getSheetName().equals(name)){
				return sheet;
			}
		}
		return null;
	}
	
	protected void checkOwnership(NSheet sheet){
		if(!sheets.contains(sheet)){
			throw new InvalidateModelOpException("doesn't has ownersheep "+ sheet);
		}
	}
	
	protected String suggestSheetName(String basename){
		int i = 1;
		HashSet<String> names = new HashSet<String>();
		for(NSheet sheet:sheets){
			names.add(sheet.getSheetName());
		}
		String name;
		do{
			name = basename + " "+i++;
		}while(names.contains(basename));
		return name;
	}
	
	protected void sendEvent(String name,Object... data){
		
		Map<String,Object> datamap = new HashMap<String,Object>();
		datamap.put("book", this);
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
		
		//implicitly deliver to sheet
		for(NSheet sheet:sheets){
			((SheetImpl)sheet).onModelEvent(event);
		}
		//TODO implicitly deliver to book series member?
		
		//TODO post to out side listener
	}

	public NSheet createSheet(NSheet src) {
		
		if(src!=null)
			checkOwnership(src);
		
		String name = src==null?"Sheet":src.getSheetName();
		if(src!=null && !name.startsWith("Copy_")){
			name = "Copy_"+name;
		}
		name = suggestSheetName(name);
		SheetImpl sheet = new SheetImpl(this);
		if(src instanceof SheetImpl){
			((SheetImpl)src).cloneSheet((SheetImpl)sheet);
		}
		((SheetImpl)sheet).setSheetName(name);
		sheets.add(sheet);
		
		sendEvent(ModelEvents.ON_SHEET_ADDED, "sheet", sheet);
		return sheet;
	}

	public void setSheetName(NSheet sheet, String newname) {
		checkOwnership(sheet);
		if(getSheetByName(newname)!=null){
			throw new InvalidateModelOpException("sheet name "+newname+" is dpulicated");
		}
		checkLegalName(newname);
		
		String oldname = sheet.getSheetName();
		((SheetImpl)sheet).setSheetName(newname);
		
		sendEvent(ModelEvents.ON_SHEET_RENAMED, "sheet", sheet, "oldName", oldname);
	}

	protected void checkLegalName(String newname) {
		//TODO
	}

	public void deleteSheet(NSheet sheet) {
		checkOwnership(sheet);
		if(sheets.size()==1){
			throw new InvalidateModelOpException("can't remove last sheet");
		}
		int index = sheets.indexOf(sheet);
		sheets.remove(index);
		
		sendEvent(ModelEvents.ON_SHEET_DELETED, "sheet", sheet, "index", index);
	}

	public void moveSheetTo(NSheet sheet, int index) {
		checkOwnership(sheet);
		if(index<0|| index>=sheets.size()){
			throw new InvalidateModelOpException("new sheet position out of bound "+sheets.size() +"," +index);
		}
		int oldindex = sheets.indexOf(sheet);
		if(oldindex==index){
			return;
		}
		sheets.remove(oldindex);
		sheets.add(index, (SheetImpl)sheet);
		sendEvent(ModelEvents.ON_SHEET_DELETED, "sheet", sheet, "index", index, "oldIndex", oldindex);
	}

	public void dump(StringBuilder builder) {
		for(SheetImpl sheet:sheets){
			builder.append("==Sheet==\n");
			builder.append(sheet.getSheetName()).append("[\n");
			sheet.dump(builder);
			builder.append("]\n");
		}
	}
}
