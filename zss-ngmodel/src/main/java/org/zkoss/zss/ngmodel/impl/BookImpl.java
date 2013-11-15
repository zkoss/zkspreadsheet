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
import org.zkoss.zss.ngmodel.util.Strings;

/**
 * @author dennis
 *
 */
public class BookImpl implements NBook{
	
	protected List<SheetImpl> sheets = new LinkedList<SheetImpl>();
	
	
	public BookImpl(){
	}
	
	public NSheet getSheet(int i){
		return sheets.get(i);
	}
	
	public int getNumOfSheet(){
		return sheets.size();
	}
	
	public NSheet getSheetByName(String name){
		for(NSheet sheet:sheets){
//			if(sheet.getSheetName().equals(name)){
			if(sheet.getSheetName().equalsIgnoreCase(name)){
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
	
//	protected String suggestSheetName(String basename){
//		int i = 1;
//		HashSet<String> names = new HashSet<String>();
//		for(NSheet sheet:sheets){
//			names.add(sheet.getSheetName());
//		}
//		String name = basename==null?"Sheet 1":basename;
//		while(names.contains(name)){
//			name = basename + " "+i++;
//		};
//		return name;
//	}
	
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
	
	public NSheet createSheet(String name) {
		return createSheet(name,null);
	}
	public NSheet createSheet(String name,NSheet src) {
		checkLegalName(name);
		if(src!=null)
			checkOwnership(src);
		

		SheetImpl sheet = new SheetImpl(this);
		if(src instanceof SheetImpl){
			((SheetImpl)src).copySheet((SheetImpl)sheet);
		}
		((SheetImpl)sheet).setSheetName(name);
		sheets.add(sheet);
		
		sendEvent(ModelEvents.ON_SHEET_ADDED, "sheet", sheet);
		return sheet;
	}

	public void setSheetName(NSheet sheet, String newname) {
		checkLegalName(newname);
		checkOwnership(sheet);
		
		String oldname = sheet.getSheetName();
		((SheetImpl)sheet).setSheetName(newname);
		
		sendEvent(ModelEvents.ON_SHEET_RENAMED, "sheet", sheet, "oldName", oldname);
	}

	protected void checkLegalName(String name) {
		if(Strings.isBlank(name)){
			throw new InvalidateModelOpException("sheet name '"+name+"' is not legal");
		}
		if(getSheetByName(name)!=null){
			throw new InvalidateModelOpException("sheet name '"+name+"' is dpulicated");
		}
		
		//TODO
	}

	public void deleteSheet(NSheet sheet) {
		checkOwnership(sheet);

		int index = sheets.indexOf(sheet);
		sheets.remove(index);
		
		sendEvent(ModelEvents.ON_SHEET_DELETED, "sheet", sheet, "index", index);
	}

	public void moveSheetTo(NSheet sheet, int index) {
		checkOwnership(sheet);
		if(index<0|| index>=sheets.size()){
			throw new InvalidateModelOpException("new position out of bound "+sheets.size() +"<>" +index);
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
			sheet.dump(builder);
		}
	}
}
