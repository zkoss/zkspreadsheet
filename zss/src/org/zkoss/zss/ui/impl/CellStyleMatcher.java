/* StyleMatcher.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/10/16 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */
package org.zkoss.zss.ui.impl;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.zss.model.impl.BookHelper;

/**
 * @author dennis
 * 
 */
public class CellStyleMatcher {

	
	Map<Property,Object> criteria = new LinkedHashMap<Property, Object>();
	
	private enum Property {
		Alignment,
		VerticalAlignment,
		
		BorderBottom,
		BorderRight,
		BorderTop,
		BorderLeft,
		BottomBorderColor,
		RightBorderColor,
		TopBorderColor,
		LeftBorderColor,
		
		DataFormat,
		
		FillBackgroundColor,
		FillForegroundColor,
		FillPattern,
		
		FontIndex,
		
		Hidden,
		Indention,
		Locked,
		Rotation,

		WrapText;
	}

	
	public CellStyleMatcher(){}
	public CellStyleMatcher(Workbook book,CellStyle criteria){
		setAlignment(criteria.getAlignment());
		setVerticalAlignment(criteria.getVerticalAlignment());
		short btype;
		setBorderBottom(btype=criteria.getBorderBottom());
		if(btype!=CellStyle.BORDER_NONE){//only compare color when the border is not none
			setBottomBorderColor(BookHelper.colorToBorderHTML(book,criteria.getBottomBorderColorColor()));
		}
		setBorderLeft(btype=criteria.getBorderLeft());
		if(btype!=CellStyle.BORDER_NONE){
			setLeftBorderColor(BookHelper.colorToBorderHTML(book,criteria.getLeftBorderColorColor()));
		}
		setBorderRight(btype=criteria.getBorderRight());
		if(btype!=CellStyle.BORDER_NONE){
			setRightBorderColor(BookHelper.colorToBorderHTML(book,criteria.getRightBorderColorColor()));
		}
		setBorderTop(btype=criteria.getBorderTop());
		if(btype!=CellStyle.BORDER_NONE){
			setTopBorderColor(BookHelper.colorToBorderHTML(book,criteria.getTopBorderColorColor()));
		}
		
		setDataFormat(criteria.getDataFormat());
		
		setFillBackgroundColor(BookHelper.colorToBackgroundHTML(book,criteria.getFillBackgroundColorColor()));
		setFillForegroundColor(BookHelper.colorToForegroundHTML(book,criteria.getFillForegroundColorColor()));
		setFillPattern(criteria.getFillPattern());
		
		setFontIndex(criteria.getFontIndex());
		setHidden(criteria.getHidden());
		setIndention(criteria.getIndention());
		
		setLocked(criteria.getLocked());
		
		setRotation(criteria.getRotation());
		
		setWrapText(criteria.getWrapText());
	}
	
	public void setDataFormat(short fmt) {
		criteria.put(Property.DataFormat, fmt);
	}

	public void setFontIndex(short fontIndex) {
		criteria.put(Property.FontIndex, fontIndex);
	}

	public void setHidden(boolean hidden) {
		criteria.put(Property.Hidden, hidden);
	}

	public void setLocked(boolean locked) {
		criteria.put(Property.Locked, locked);
	}

	public void setAlignment(short align) {
		criteria.put(Property.Alignment, align);
	}

	public void setWrapText(boolean wrapped) {
		criteria.put(Property.WrapText, wrapped);
	}

	public void setVerticalAlignment(short align) {
		criteria.put(Property.VerticalAlignment, align);
	}

	public void setRotation(short rotation) {
		criteria.put(Property.Rotation,rotation );
	}

	public void setIndention(short indent) {
		criteria.put(Property.Indention, indent);
	}

	public void setBorderRight(short border) {
		criteria.put(Property.BorderRight, border);
	}

	public void setBorderTop(short border) {
		criteria.put(Property.BorderTop,border );
	}

	public void setBorderBottom(short border) {
		criteria.put(Property.BorderBottom, border);
	}
	public void setBorderLeft(short border) {
		criteria.put(Property.BorderLeft, border);
	}

	public void setLeftBorderColor(String htmlcolor) {
		criteria.put(Property.LeftBorderColor,htmlcolor );
	}

	public void setRightBorderColor(String htmlcolor) {
		criteria.put(Property.RightBorderColor, htmlcolor);
	}

	public void setTopBorderColor(String htmlcolor) {
		criteria.put(Property.TopBorderColor,htmlcolor);
	}

	public void setBottomBorderColor(String htmlcolor) {
		criteria.put(Property.BottomBorderColor,htmlcolor );
	}

	public void setFillPattern(short fp) {
		criteria.put(Property.FillPattern,fp );
	}

	public void setFillBackgroundColor(String htmlcolor) {
		criteria.put(Property.FillBackgroundColor, htmlcolor);
	}

	public void setFillForegroundColor(String htmlcolor) {
		criteria.put(Property.FillForegroundColor, htmlcolor );
	}
	//remove api
	public void removeDataFormat() {
		criteria.remove(Property.DataFormat);
	}

	public void removeFontIndex() {
		criteria.remove(Property.FontIndex);
	}

	public void removeHidden() {
		criteria.remove(Property.Hidden);
	}

	public void removeLocked(){
		criteria.remove(Property.Locked);
	}

	public void removeAlignment(){
		criteria.remove(Property.Alignment);
	}

	public void removeWrapText(){
		criteria.remove(Property.WrapText);
	}

	public void removeVerticalAlignment(){
		criteria.remove(Property.VerticalAlignment);
	}

	public void removeRotation(){
		criteria.remove(Property.Rotation);
	}

	public void removeIndention(){
		criteria.remove(Property.Indention);
	}

	public void removeBorderRight(){
		criteria.remove(Property.BorderRight);
	}

	public void removeBorderTop(){
		criteria.remove(Property.BorderTop);
	}

	public void removeBorderBottom(){
		criteria.remove(Property.BorderBottom);
	}
	public void removeBorderLeft(){
		criteria.remove(Property.BorderLeft);
	}

	public void removeLeftBorderColor(){
		criteria.remove(Property.LeftBorderColor);
	}

	public void removeRightBorderColor(){
		criteria.remove(Property.RightBorderColor);
	}

	public void removeTopBorderColor(){
		criteria.remove(Property.TopBorderColor);
	}

	public void removeBottomBorderColor(){
		criteria.remove(Property.BottomBorderColor);
	}

	public void removeFillPattern(){
		criteria.remove(Property.FillPattern);
	}

	public void removeFillBackgroundColor(){
		criteria.remove(Property.FillBackgroundColor);
	}

	public void removeFillForegroundColor(){
		criteria.remove(Property.FillForegroundColor);
	}
	public boolean match(Workbook book,CellStyle style){
		for(Entry<Property,Object> e:criteria.entrySet()){
			switch(e.getKey()){
			case Alignment:
				if(!equals(e.getValue(),style.getAlignment())){
					return false;
				}
				break;
			case VerticalAlignment:
				if(!equals(e.getValue(),style.getVerticalAlignment())){
					return false;
				}
				break;
			case BorderBottom:
				if(!equals(e.getValue(),style.getBorderBottom())){
					return false;
				}
				break;
			case BorderRight:
				if(!equals(e.getValue(),style.getBorderRight())){
					return false;
				}
				break;
			case BorderTop:
				if(!equals(e.getValue(),style.getBorderTop())){
					return false;
				}
				break;
			case BorderLeft:
				if(!equals(e.getValue(),style.getBorderLeft())){
					return false;
				}
				break;
			case BottomBorderColor:
				if(!equals(e.getValue(),BookHelper.colorToBorderHTML(book, style.getBottomBorderColorColor()))){
					return false;
				}
				break;
			case RightBorderColor:
				if(!equals(e.getValue(),BookHelper.colorToBorderHTML(book, style.getRightBorderColorColor()))){
					return false;
				}
				break;
			case TopBorderColor:
				if(!equals(e.getValue(),BookHelper.colorToBorderHTML(book, style.getTopBorderColorColor()))){
					return false;
				}
				break;
			case LeftBorderColor:
				if(!equals(e.getValue(),BookHelper.colorToBorderHTML(book, style.getLeftBorderColorColor()))){
					return false;
				}
				break;
			case DataFormat:
				if(!equals(e.getValue(),style.getDataFormat())){
					return false;
				}
				break;
			case FillBackgroundColor:
				if(!equals(e.getValue(),BookHelper.colorToBackgroundHTML(book, style.getFillBackgroundColorColor()))){
					return false;
				}
				break;
			case FillForegroundColor:
				if(!equals(e.getValue(),BookHelper.colorToForegroundHTML(book, style.getFillForegroundColorColor()))){
					return false;
				}
				break;
			case FillPattern:
				if(!equals(e.getValue(),style.getFillPattern())){
					return false;
				}
				break;			
			case FontIndex:
				if(!equals(e.getValue(),style.getFontIndex())){
					return false;
				}
				break;
			case Hidden:
				if(!equals(e.getValue(),style.getHidden())){
					return false;
				}
				break;
			case Indention:
				if(!equals(e.getValue(),style.getIndention())){
					return false;
				}
				break;
			case Locked:
				if(!equals(e.getValue(),style.getLocked())){
					return false;
				}
				break;
			case Rotation:
				if(!equals(e.getValue(),style.getRotation())){
					return false;
				}
				break;
			case WrapText:
				if(!equals(e.getValue(),style.getWrapText())){
					return false;
				}
				break;
			}
		}
		return true;
	}
	
	public boolean equals(Object o1,Object o2){
		if(o1==o2)
			return true;
		return o1!=null?o1.equals(o2):false;
	}
}
