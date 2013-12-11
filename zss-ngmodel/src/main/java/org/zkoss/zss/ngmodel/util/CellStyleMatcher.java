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

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NFont;


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
		BorderBottomColor,
		BorderRightColor,
		BorderTopColor,
		BorderLeftColor,
		
		DataFormat,
		
		FillColor,
//		FillForegroundColor,
		FillPattern,
		
		FontName,
		FontColor,
		FontBoldweight,
		FontHeightPoints,
		FontItalic,
		FontStrikeout,
		FontTypeOffset,
		FontUnderline,
		
		Hidden,
//		Indention,
		Locked,
//		Rotation,

		WrapText;
	}

	
	public CellStyleMatcher(){}
	public CellStyleMatcher(NCellStyle criteria){
		setAlignment(criteria.getAlignment());
		setVerticalAlignment(criteria.getVerticalAlignment());
		NCellStyle.BorderType btype;
		setBorderBottom(btype=criteria.getBorderBottom());
		if(btype!=NCellStyle.BorderType.NONE){//only compare color when the border is not none
			setBorderBottomColor(criteria.getBorderBottomColor().getHtmlColor());
		}
		setBorderLeft(btype=criteria.getBorderLeft());
		if(btype!=NCellStyle.BorderType.NONE){
			setBorderLeftColor(criteria.getBorderLeftColor().getHtmlColor());
		}
		setBorderRight(btype=criteria.getBorderRight());
		if(btype!=NCellStyle.BorderType.NONE){
			setBorderRightColor(criteria.getBorderRightColor().getHtmlColor());
		}
		setBorderTop(btype=criteria.getBorderTop());
		if(btype!=NCellStyle.BorderType.NONE){
			setBorderTopColor(criteria.getBorderTopColor().getHtmlColor());
		}
		
		setDataFormat(criteria.getDataFormat());
		
		setFillColor(criteria.getFillColor().getHtmlColor());
//		setFillForegroundColor(BookHelper.colorToForegroundHTML(book,criteria.getFillForegroundColorColor()));
		setFillPattern(criteria.getFillPattern());
		
		setFont(criteria.getFont());
		
		
		setHidden(criteria.isHidden());
//		setIndention(criteria.getIndention());
		
		setLocked(criteria.isLocked());
		
//		setRotation(criteria.getRotation());
		
		setWrapText(criteria.isWrapText());
	}
	
	public void setDataFormat(String fmt) {
		criteria.put(Property.DataFormat, fmt);
	}

	public void setFontColor(String color) {
		criteria.put(Property.FontColor, color);
	}
	
	public void setFontName(String name) {
		criteria.put(Property.FontName, name);
	}
	
	public void setFontBoldweight(NFont.Boldweight boldweight) {
		criteria.put(Property.FontBoldweight, boldweight);
	}
	
	public void setFontHeightPoints(int height) {
		criteria.put(Property.FontHeightPoints, height);
	}
	
	public void setFontItalic(boolean italic) {
		criteria.put(Property.FontItalic, italic);
	}
	
	public void setFontStrikeout(boolean strikeout) {
		criteria.put(Property.FontStrikeout, strikeout);
	}
	
	public void setFontTypeOffset(NFont.TypeOffset typeOffset) {
		criteria.put(Property.FontTypeOffset, typeOffset);
	}
	
	public void setFontUnderline(NFont.Underline underline) {
		criteria.put(Property.FontUnderline, underline);
	}
	
	public void setFont(NFont font) {
		setFontColor(font.getColor().getHtmlColor());
		setFontName(font.getName());
		setFontBoldweight(font.getBoldweight());
		setFontHeightPoints(font.getHeightPoints());
		setFontItalic(font.isItalic());
		setFontStrikeout(font.isStrikeout());
		setFontTypeOffset(font.getTypeOffset());
		setFontUnderline(font.getUnderline());
	}
	public void removeFont() {
		removeFontColor();
		removeFontName();
		removeFontBoldweight();
		removeFontHeightPoints();
		removeFontItalic();
		removeFontStrikeout();
		removeFontTypeOffset();
		removeFontUnderline();
	}

	public void setHidden(boolean hidden) {
		criteria.put(Property.Hidden, hidden);
	}

	public void setLocked(boolean locked) {
		criteria.put(Property.Locked, locked);
	}

	public void setAlignment(NCellStyle.Alignment align) {
		criteria.put(Property.Alignment, align);
	}

	public void setWrapText(boolean wrapped) {
		criteria.put(Property.WrapText, wrapped);
	}

	public void setVerticalAlignment(NCellStyle.VerticalAlignment align) {
		criteria.put(Property.VerticalAlignment, align);
	}

//	public void setRotation(short rotation) {
//		criteria.put(Property.Rotation,rotation );
//	}
//
//	public void setIndention(short indent) {
//		criteria.put(Property.Indention, indent);
//	}

	public void setBorderRight(NCellStyle.BorderType border) {
		criteria.put(Property.BorderRight, border);
	}

	public void setBorderTop(NCellStyle.BorderType border) {
		criteria.put(Property.BorderTop,border );
	}

	public void setBorderBottom(NCellStyle.BorderType border) {
		criteria.put(Property.BorderBottom, border);
	}
	public void setBorderLeft(NCellStyle.BorderType border) {
		criteria.put(Property.BorderLeft, border);
	}

	public void setBorderLeftColor(String htmlcolor) {
		criteria.put(Property.BorderLeftColor,htmlcolor );
	}

	public void setBorderRightColor(String htmlcolor) {
		criteria.put(Property.BorderRightColor, htmlcolor);
	}

	public void setBorderTopColor(String htmlcolor) {
		criteria.put(Property.BorderTopColor,htmlcolor);
	}

	public void setBorderBottomColor(String htmlcolor) {
		criteria.put(Property.BorderBottomColor,htmlcolor );
	}

	public void setFillPattern(NCellStyle.FillPattern fp) {
		criteria.put(Property.FillPattern,fp );
	}

	public void setFillColor(String htmlcolor) {
		criteria.put(Property.FillColor, htmlcolor);
	}

	//remove api
	public void removeDataFormat() {
		criteria.remove(Property.DataFormat);
	}

	public void removeFontColor() {
		criteria.remove(Property.FontColor);
	}
	public void removeFontName() {
		criteria.remove(Property.FontName);
	}
	public void removeFontBoldweight() {
		criteria.remove(Property.FontBoldweight);
	}
	public void removeFontHeightPoints() {
		criteria.remove(Property.FontHeightPoints);
	}
	public void removeFontItalic() {
		criteria.remove(Property.FontItalic);
	}
	public void removeFontStrikeout() {
		criteria.remove(Property.FontStrikeout);
	}
	public void removeFontTypeOffset() {
		criteria.remove(Property.FontTypeOffset);
	}
	public void removeFontUnderline() {
		criteria.remove(Property.FontUnderline);
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

//	public void removeRotation(){
//		criteria.remove(Property.Rotation);
//	}
//
//	public void removeIndention(){
//		criteria.remove(Property.Indention);
//	}

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

	public void removeBorderLeftColor(){
		criteria.remove(Property.BorderLeftColor);
	}

	public void removeBorderRightColor(){
		criteria.remove(Property.BorderRightColor);
	}

	public void removeBorderTopColor(){
		criteria.remove(Property.BorderTopColor);
	}

	public void removeBorderBottomColor(){
		criteria.remove(Property.BorderBottomColor);
	}

	public void removeFillPattern(){
		criteria.remove(Property.FillPattern);
	}

	public void removeFillColor(){
		criteria.remove(Property.FillColor);
	}

//	public void removeFillForegroundColor(){
//		criteria.remove(Property.FillForegroundColor);
//	}
	
	public boolean match(NCellStyle style){
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
			case BorderBottomColor:
				if(!equals(e.getValue(),style.getBorderBottomColor().getHtmlColor())){
					return false;
				}
				break;
			case BorderRightColor:
				if(!equals(e.getValue(),style.getBorderRightColor().getHtmlColor())){
					return false;
				}
				break;
			case BorderTopColor:
				if(!equals(e.getValue(),style.getBorderTopColor().getHtmlColor())){
					return false;
				}
				break;
			case BorderLeftColor:
				if(!equals(e.getValue(),style.getBorderLeftColor().getHtmlColor())){
					return false;
				}
				break;
			case DataFormat:
				if(!equals(e.getValue(),style.getDataFormat())){
					return false;
				}
				break;
			case FillColor:
				if(!equals(e.getValue(),style.getFillColor().getHtmlColor())){
					return false;
				}
				break;
//			case FillForegroundColor:
//				if(!equals(e.getValue(),BookHelper.colorToForegroundHTML(book, style.getFillForegroundColorColor()))){
//					return false;
//				}
//				break;
			case FillPattern:
				if(!equals(e.getValue(),style.getFillPattern())){
					return false;
				}
				break;			
			case FontColor:
				if(!equals(e.getValue(),style.getFont().getColor().getHtmlColor())){
					return false;
				}
				break;
			case FontName:
				if(!equals(e.getValue(),style.getFont().getName())){
					return false;
				}
				break;
			case FontBoldweight:
				if(!equals(e.getValue(),style.getFont().getBoldweight())){
					return false;
				}
				break;
			case FontHeightPoints:
				if(!equals(e.getValue(),style.getFont().getHeightPoints())){
					return false;
				}
				break;
			case FontItalic:
				if(!equals(e.getValue(),style.getFont().isItalic())){
					return false;
				}
				break;
			case FontStrikeout:
				if(!equals(e.getValue(),style.getFont().isStrikeout())){
					return false;
				}
				break;
			case FontTypeOffset:
				if(!equals(e.getValue(),style.getFont().getTypeOffset())){
					return false;
				}
				break;
			case FontUnderline:
				if(!equals(e.getValue(),style.getFont().getUnderline())){
					return false;
				}
				break;
				
			case Hidden:
				if(!equals(e.getValue(),style.isHidden())){
					return false;
				}
				break;
//			case Indention:
//				if(!equals(e.getValue(),style.getIndention())){
//					return false;
//				}
//				break;
			case Locked:
				if(!equals(e.getValue(),style.isLocked())){
					return false;
				}
				break;
//			case Rotation:
//				if(!equals(e.getValue(),style.getRotation())){
//					return false;
//				}
//				break;
			case WrapText:
				if(!equals(e.getValue(),style.isWrapText())){
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
