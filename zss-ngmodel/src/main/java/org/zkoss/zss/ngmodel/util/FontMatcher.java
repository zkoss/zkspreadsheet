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

package org.zkoss.zss.ngmodel.util;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.zkoss.zss.ngmodel.NFont;


/**
 * @author dennis
 * @since 3.5.0
 */
public class FontMatcher {

	
	Map<Property,Object> criteria = new LinkedHashMap<Property, Object>();
	
	private enum Property {
		
		Name,
		Color,
		Boldweight,
		HeightPoints,
		Italic,
		Strikeout,
		TypeOffset,
		Underline,
	}

	
	public FontMatcher(){}
	public FontMatcher(NFont criteria){
		setColor(criteria.getColor().getHtmlColor());
		setName(criteria.getName());
		setBoldweight(criteria.getBoldweight());
		setHeightPoints(criteria.getHeightPoints());
		setItalic(criteria.isItalic());
		setStrikeout(criteria.isStrikeout());
		setTypeOffset(criteria.getTypeOffset());
		setUnderline(criteria.getUnderline());
		
	}
	
	public void setColor(String color) {
		criteria.put(Property.Color, color);
	}
	
	public void setName(String name) {
		criteria.put(Property.Name, name);
	}
	
	public void setBoldweight(NFont.Boldweight boldweight) {
		criteria.put(Property.Boldweight, boldweight);
	}
	
	public void setHeightPoints(int height) {
		criteria.put(Property.HeightPoints, height);
	}
	
	public void setItalic(boolean italic) {
		criteria.put(Property.Italic, italic);
	}
	
	public void setStrikeout(boolean strikeout) {
		criteria.put(Property.Strikeout, strikeout);
	}
	
	public void setTypeOffset(NFont.TypeOffset typeOffset) {
		criteria.put(Property.TypeOffset, typeOffset);
	}
	
	public void setUnderline(NFont.Underline underline) {
		criteria.put(Property.Underline, underline);
	}

	//remove api

	public void removeColor() {
		criteria.remove(Property.Color);
	}
	public void removeName() {
		criteria.remove(Property.Name);
	}
	public void removeBoldweight() {
		criteria.remove(Property.Boldweight);
	}
	public void removeHeightPoints() {
		criteria.remove(Property.HeightPoints);
	}
	public void removeItalic() {
		criteria.remove(Property.Italic);
	}
	public void removeStrikeout() {
		criteria.remove(Property.Strikeout);
	}
	public void removeTypeOffset() {
		criteria.remove(Property.TypeOffset);
	}
	public void removeUnderline() {
		criteria.remove(Property.Underline);
	}

	
	
	public boolean match(NFont style){
		for(Entry<Property,Object> e:criteria.entrySet()){
			switch(e.getKey()){
			case Color:
				if(!equals(e.getValue(),style.getColor().getHtmlColor())){
					return false;
				}
				break;
			case Name:
				if(!equals(e.getValue(),style.getName())){
					return false;
				}
				break;
			case Boldweight:
				if(!equals(e.getValue(),style.getBoldweight())){
					return false;
				}
				break;
			case HeightPoints:
				if(!equals(e.getValue(),style.getHeightPoints())){
					return false;
				}
				break;
			case Italic:
				if(!equals(e.getValue(),style.isItalic())){
					return false;
				}
				break;
			case Strikeout:
				if(!equals(e.getValue(),style.isStrikeout())){
					return false;
				}
				break;
			case TypeOffset:
				if(!equals(e.getValue(),style.getTypeOffset())){
					return false;
				}
				break;
			case Underline:
				if(!equals(e.getValue(),style.getUnderline())){
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
