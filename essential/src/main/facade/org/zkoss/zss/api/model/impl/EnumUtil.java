package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.NRange.PasteOperation;
import org.zkoss.zss.api.NRange.PasteType;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.TypeOffset;
import org.zkoss.zss.api.model.NFont.Underline;
import org.zkoss.zss.model.Range;

public class EnumUtil {
	
	private static <T> void assertArgNotNull(T obj,String name){
		if(obj == null){
			throw new IllegalArgumentException("argument "+name==null?"":name+" is null");
		}
	}
	
	public static int toRangePasteOpNative(PasteOperation op) {
		assertArgNotNull(op,"paste operation");
		switch(op){
		case PASTEOP_ADD:
			return Range.PASTEOP_ADD;
		case PASTEOP_SUB:
			return Range.PASTEOP_SUB;
		case PASTEOP_MUL:
			return Range.PASTEOP_MUL;
		case PASTEOP_DIV:
			return Range.PASTEOP_DIV;
		case PASTEOP_NONE:
			return Range.PASTEOP_NONE;
		}
		throw new IllegalArgumentException("unknow paste operation "+op);
	}


	public static int toRangePasteTypeNative(PasteType type) {
		assertArgNotNull(type,"paste type");
		switch(type){
		case PASTE_ALL:
			return Range.PASTE_ALL;
		case PASTE_ALL_EXCEPT_BORDERS:
			return Range.PASTE_ALL_EXCEPT_BORDERS;
		case PASTE_COLUMN_WIDTHS:
			return Range.PASTE_COLUMN_WIDTHS;
		case PASTE_COMMENTS:
			return Range.PASTE_COMMENTS;
		case PASTE_FORMATS:
			return Range.PASTE_FORMATS;
		case PASTE_FORMULAS:
			return Range.PASTE_FORMULAS;
		case PASTE_FORMULAS_AND_NUMBER_FORMATS:
			return Range.PASTE_FORMULAS_AND_NUMBER_FORMATS;
		case PASTE_VALIDATAION:
			return Range.PASTE_VALIDATAION;
		case PASTE_VALUES:
			return Range.PASTE_VALUES;
		case PASTE_VALUES_AND_NUMBER_FORMATS:
			return Range.PASTE_VALUES_AND_NUMBER_FORMATS;
		}
		throw new IllegalArgumentException("unknow paste operation "+type);
	}
	
	public static TypeOffset toFontTypeOffset(short typeOffset){
		switch(typeOffset){
		case Font.SS_NONE:
			return NFont.TypeOffset.NONE;
		case Font.SS_SUB:
			return NFont.TypeOffset.SUB;
		case Font.SS_SUPER:
			return NFont.TypeOffset.SUPER;
		}
		throw new IllegalArgumentException("unknow font type offset "+typeOffset);
	}

	public static short toFontTypeOffset(TypeOffset typeOffset) {
		assertArgNotNull(typeOffset,"typeOffset");
		switch(typeOffset){
		case NONE:
			return Font.SS_NONE;
		case SUB:
			return Font.SS_SUB;
		case SUPER:
			return Font.SS_SUPER;
		}
		throw new IllegalArgumentException("unknow font type offset "+typeOffset);
	}

	public static Underline toFontUnderline(byte underline) {
		switch(underline){
		case Font.U_NONE:
			return NFont.Underline.NONE;
		case Font.U_SINGLE:
			return NFont.Underline.SINGLE;
		case Font.U_SINGLE_ACCOUNTING:
			return NFont.Underline.SINGLE_ACCOUNTING;
		case Font.U_DOUBLE:
			return NFont.Underline.DOUBLE;
		case Font.U_DOUBLE_ACCOUNTING:
			return NFont.Underline.DOUBLE_ACCOUNTING;
		}
		throw new IllegalArgumentException("unknow font underline "+underline);
	}


	public static byte toFontUnderline(Underline underline) {
		assertArgNotNull(underline,"underline");
		switch(underline){
		case NONE:
			return Font.U_NONE;
		case SINGLE:
			return Font.U_SINGLE;
		case SINGLE_ACCOUNTING:
			return Font.U_SINGLE_ACCOUNTING;
		case DOUBLE:
			return Font.U_DOUBLE;
		case DOUBLE_ACCOUNTING:
			return Font.U_DOUBLE_ACCOUNTING;
		}
		throw new IllegalArgumentException("unknow font underline "+underline);
	}

	public static Boldweight toFontBoldweight(short boldweight) {
		switch(boldweight){
		case Font.BOLDWEIGHT_BOLD:
			return NFont.Boldweight.BOLD;
		case Font.BOLDWEIGHT_NORMAL:
			return NFont.Boldweight.NORMAL;
		}
		throw new IllegalArgumentException("unknow font boldweight "+boldweight);
	}
	
	public static short toFontBoldweight(Boldweight boldweight) {
		switch(boldweight){
		case BOLD:
			return Font.BOLDWEIGHT_BOLD;
		case NORMAL:
			return Font.BOLDWEIGHT_NORMAL;
		}
		throw new IllegalArgumentException("unknow font boldweight "+boldweight);
	}
}
