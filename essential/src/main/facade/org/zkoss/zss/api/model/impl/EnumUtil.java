package org.zkoss.zss.api.model.impl;

import org.apache.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.NRange.PasteOperation;
import org.zkoss.zss.api.NRange.PasteType;
import org.zkoss.zss.api.model.NCellStyle.FillPatternType;
import org.zkoss.zss.api.model.NCellStyle;
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

	public static FillPatternType toStyleFillPattern(short pattern) {
		switch(pattern){
		case CellStyle.NO_FILL:
			return NCellStyle.FillPatternType.NO_FILL;
		case CellStyle.SOLID_FOREGROUND:
			return NCellStyle.FillPatternType.SOLID_FOREGROUND;
		case CellStyle.FINE_DOTS:
			return NCellStyle.FillPatternType.FINE_DOTS;
		case CellStyle.ALT_BARS:
			return NCellStyle.FillPatternType.ALT_BARS;
		case CellStyle.SPARSE_DOTS:
			return NCellStyle.FillPatternType.SPARSE_DOTS;
		case CellStyle.THICK_HORZ_BANDS:
			return NCellStyle.FillPatternType.THICK_HORZ_BANDS;
		case CellStyle.THICK_VERT_BANDS:
			return NCellStyle.FillPatternType.THICK_VERT_BANDS;
		case CellStyle.THICK_BACKWARD_DIAG:
			return NCellStyle.FillPatternType.THICK_BACKWARD_DIAG;
		case CellStyle.THICK_FORWARD_DIAG:
			return NCellStyle.FillPatternType.THICK_FORWARD_DIAG;
		case CellStyle.BIG_SPOTS:
			return NCellStyle.FillPatternType.BIG_SPOTS;
		case CellStyle.BRICKS:
			return NCellStyle.FillPatternType.BRICKS;
		case CellStyle.THIN_HORZ_BANDS:
			return NCellStyle.FillPatternType.THIN_HORZ_BANDS;
		case CellStyle.THIN_VERT_BANDS:
			return NCellStyle.FillPatternType.THIN_VERT_BANDS;
		case CellStyle.THIN_BACKWARD_DIAG:
			return NCellStyle.FillPatternType.THIN_BACKWARD_DIAG;
		case CellStyle.THIN_FORWARD_DIAG:
			return NCellStyle.FillPatternType.THIN_FORWARD_DIAG;
		case CellStyle.SQUARES:
			return NCellStyle.FillPatternType.SQUARES;
		case CellStyle.DIAMONDS:
			return NCellStyle.FillPatternType.DIAMONDS;
		case CellStyle.LESS_DOTS:
			return NCellStyle.FillPatternType.LESS_DOTS;
		case CellStyle.LEAST_DOTS:
			return NCellStyle.FillPatternType.LEAST_DOTS;
		}
		throw new IllegalArgumentException("unknow pattern type "+pattern);	}
	
	public static short toStyleFillPattern(FillPatternType pattern) {
		switch(pattern){
		case NO_FILL:
			return CellStyle.NO_FILL;
		case SOLID_FOREGROUND:
			return CellStyle.SOLID_FOREGROUND;
		case FINE_DOTS:
			return CellStyle.FINE_DOTS;
		case ALT_BARS:
			return CellStyle.ALT_BARS;
		case SPARSE_DOTS:
			return CellStyle.SPARSE_DOTS;
		case THICK_HORZ_BANDS:
			return CellStyle.THICK_HORZ_BANDS;
		case THICK_VERT_BANDS:
			return CellStyle.THICK_VERT_BANDS;
		case THICK_BACKWARD_DIAG:
			return CellStyle.THICK_BACKWARD_DIAG;
		case THICK_FORWARD_DIAG:
			return CellStyle.THICK_FORWARD_DIAG;
		case BIG_SPOTS:
			return CellStyle.BIG_SPOTS;
		case BRICKS:
			return CellStyle.BRICKS;
		case THIN_HORZ_BANDS:
			return CellStyle.THIN_HORZ_BANDS;
		case THIN_VERT_BANDS:
			return CellStyle.THIN_VERT_BANDS;
		case THIN_BACKWARD_DIAG:
			return CellStyle.THIN_BACKWARD_DIAG;
		case THIN_FORWARD_DIAG:
			return CellStyle.THIN_FORWARD_DIAG;
		case SQUARES:
			return CellStyle.SQUARES;
		case DIAMONDS:
			return CellStyle.DIAMONDS;
		case LESS_DOTS:
			return CellStyle.LESS_DOTS;
		case LEAST_DOTS:
			return CellStyle.LEAST_DOTS;
		}
		throw new IllegalArgumentException("unknow pattern type "+pattern);
	}
}
