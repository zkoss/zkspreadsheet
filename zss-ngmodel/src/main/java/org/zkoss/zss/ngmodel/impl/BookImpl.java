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
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.SpreadsheetVersion;
import org.zkoss.zss.ngmodel.util.Strings;
import org.zkoss.zss.ngmodel.util.Validations;

/**
 * @author dennis
 *
 */
public class BookImpl extends AbstractBook{
	
	protected List<AbstractSheet> sheets = new LinkedList<AbstractSheet>();
	CellStyleImpl defaultCellStyle;
	
	int sheetsId = 0;
	private int maxRowSize = SpreadsheetVersion.EXCEL2007.getMaxRows();
	private int maxColumnSize = SpreadsheetVersion.EXCEL2007.getMaxColumns();
	
	public BookImpl(){
		defaultCellStyle = new CellStyleImpl();
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
		for(AbstractSheet sheet:sheets){
			sheet.onModelEvent(event);
		}
		//TODO implicitly deliver to book series member?
		
		//TODO post to out side listener
	}
	
	public NSheet createSheet(String name) {
		return createSheet(name,null);
	}
	
	String nextSheetId(){
		StringBuilder sb = new StringBuilder("s");
		sb.append(++sheetsId);
		return sb.toString();
	}
	public NSheet createSheet(String name,NSheet src) {
		checkLegalName(name);
		if(src!=null)
			checkOwnership(src);
		

		SheetImpl sheet = new SheetImpl(this,nextSheetId());
		if(src instanceof AbstractSheet){
			((AbstractSheet)src).copyTo(sheet);
		}
		((AbstractSheet)sheet).setSheetName(name);
		sheets.add(sheet);
		
		sendEvent(ModelEvents.ON_SHEET_ADDED, "sheet", sheet);
		return sheet;
	}

	public void setSheetName(NSheet sheet, String newname) {
		checkLegalName(newname);
		checkOwnership(sheet);
		
		String oldname = sheet.getSheetName();
		((AbstractSheet)sheet).setSheetName(newname);
		
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
		
		((AbstractSheet)sheet).release();
		
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
		sheets.add(index, (AbstractSheet)sheet);
		sendEvent(ModelEvents.ON_SHEET_DELETED, "sheet", sheet, "index", index, "oldIndex", oldindex);
	}

	public void dump(StringBuilder builder) {
		for(AbstractSheet sheet:sheets){
			if(sheet instanceof SheetImpl){
				((SheetImpl)sheet).dump(builder);
			}else{
				builder.append("\n").append(sheet);
			}
		}
	}

	public NCellStyle getDefaultCellStyle() {
		return defaultCellStyle;
	}

	public NCellStyle createCellStyle() {
		return createCellStyle(null);
	}

	public NCellStyle createCellStyle(NCellStyle src) {
		if(src!=null){
			Validations.argInstance(src, AbstractCellStyle.class);
		}
		CellStyleImpl style = new CellStyleImpl();
		if(src!=null){
			((AbstractCellStyle)src).copyTo(style);
		}
		//TODO put to style table.
		
		return style;
	}

	public int getMaxRowSize() {
		return maxRowSize;
	}

	public int getMaxColumnSize() {
		return maxColumnSize;
	}
}
