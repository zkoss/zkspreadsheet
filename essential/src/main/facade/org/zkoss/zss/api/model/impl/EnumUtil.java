package org.zkoss.zss.api.model.impl;

import org.apache.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.zss.api.NRange.ApplyBorderType;
import org.zkoss.zss.api.NRange.DeleteShift;
import org.zkoss.zss.api.NRange.InsertCopyOrigin;
import org.zkoss.zss.api.NRange.InsertShift;
import org.zkoss.zss.api.NRange.PasteOperation;
import org.zkoss.zss.api.NRange.PasteType;
import org.zkoss.zss.api.model.NCellStyle.Alignment;
import org.zkoss.zss.api.model.NCellStyle.BorderType;
import org.zkoss.zss.api.model.NCellStyle.FillPattern;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NCellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.NFont;
import org.zkoss.zss.api.model.NFont.Boldweight;
import org.zkoss.zss.api.model.NFont.TypeOffset;
import org.zkoss.zss.api.model.NFont.Underline;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.impl.BookHelper;

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

	public static FillPattern toStyleFillPattern(short pattern) {
		switch(pattern){
		case CellStyle.NO_FILL:
			return NCellStyle.FillPattern.NO_FILL;
		case CellStyle.SOLID_FOREGROUND:
			return NCellStyle.FillPattern.SOLID_FOREGROUND;
		case CellStyle.FINE_DOTS:
			return NCellStyle.FillPattern.FINE_DOTS;
		case CellStyle.ALT_BARS:
			return NCellStyle.FillPattern.ALT_BARS;
		case CellStyle.SPARSE_DOTS:
			return NCellStyle.FillPattern.SPARSE_DOTS;
		case CellStyle.THICK_HORZ_BANDS:
			return NCellStyle.FillPattern.THICK_HORZ_BANDS;
		case CellStyle.THICK_VERT_BANDS:
			return NCellStyle.FillPattern.THICK_VERT_BANDS;
		case CellStyle.THICK_BACKWARD_DIAG:
			return NCellStyle.FillPattern.THICK_BACKWARD_DIAG;
		case CellStyle.THICK_FORWARD_DIAG:
			return NCellStyle.FillPattern.THICK_FORWARD_DIAG;
		case CellStyle.BIG_SPOTS:
			return NCellStyle.FillPattern.BIG_SPOTS;
		case CellStyle.BRICKS:
			return NCellStyle.FillPattern.BRICKS;
		case CellStyle.THIN_HORZ_BANDS:
			return NCellStyle.FillPattern.THIN_HORZ_BANDS;
		case CellStyle.THIN_VERT_BANDS:
			return NCellStyle.FillPattern.THIN_VERT_BANDS;
		case CellStyle.THIN_BACKWARD_DIAG:
			return NCellStyle.FillPattern.THIN_BACKWARD_DIAG;
		case CellStyle.THIN_FORWARD_DIAG:
			return NCellStyle.FillPattern.THIN_FORWARD_DIAG;
		case CellStyle.SQUARES:
			return NCellStyle.FillPattern.SQUARES;
		case CellStyle.DIAMONDS:
			return NCellStyle.FillPattern.DIAMONDS;
		case CellStyle.LESS_DOTS:
			return NCellStyle.FillPattern.LESS_DOTS;
		case CellStyle.LEAST_DOTS:
			return NCellStyle.FillPattern.LEAST_DOTS;
		}
		throw new IllegalArgumentException("unknow pattern type "+pattern);	}
	
	public static short toStyleFillPattern(FillPattern pattern) {
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

	public static short toStyleAlignemnt(Alignment alignment) {
		switch(alignment){
		case GENERAL:
			return CellStyle.ALIGN_GENERAL;
		case LEFT:
			return CellStyle.ALIGN_LEFT;
		case CENTER:
			return CellStyle.ALIGN_CENTER;
		case RIGHT:
			return CellStyle.ALIGN_RIGHT;
		case FILL:
			return CellStyle.ALIGN_FILL;
		case JUSTIFY:
			return CellStyle.ALIGN_JUSTIFY;
		case CENTER_SELECTION:
			return CellStyle.ALIGN_CENTER_SELECTION;
		}
		throw new IllegalArgumentException("unknow cell alignment "+alignment);
	}
	public static Alignment toStyleAlignemnt(short alignment) {
		switch(alignment){
		case CellStyle.ALIGN_GENERAL:
			return Alignment.GENERAL;
		case CellStyle.ALIGN_LEFT:
			return Alignment.LEFT;
		case CellStyle.ALIGN_CENTER:
			return Alignment.CENTER;
		case CellStyle.ALIGN_RIGHT:
			return Alignment.RIGHT;
		case CellStyle.ALIGN_FILL:
			return Alignment.FILL;
		case CellStyle.ALIGN_JUSTIFY:
			return Alignment.JUSTIFY;
		case CellStyle.ALIGN_CENTER_SELECTION:
			return Alignment.CENTER_SELECTION;
		}
		throw new IllegalArgumentException("unknow cell alignment "+alignment);
	}
	public static short toStyleVerticalAlignemnt(VerticalAlignment alignment) {
		switch(alignment){
		case TOP:
			return CellStyle.VERTICAL_TOP;
		case CENTER:
			return CellStyle.VERTICAL_CENTER;
		case BOTTOM:
			return CellStyle.VERTICAL_BOTTOM;
		case JUSTIFY:
			return CellStyle.VERTICAL_JUSTIFY;
		}
		throw new IllegalArgumentException("unknow cell vertical alignment "+alignment);
	}
	public static VerticalAlignment toStyleVerticalAlignemnt(short alignment) {
		switch(alignment){
		case CellStyle.VERTICAL_TOP:
			return VerticalAlignment.TOP;
		case CellStyle.VERTICAL_CENTER:
			return VerticalAlignment.CENTER;
		case CellStyle.VERTICAL_BOTTOM:
			return VerticalAlignment.BOTTOM;
		case CellStyle.VERTICAL_JUSTIFY:
			return VerticalAlignment.JUSTIFY;
		}
		throw new IllegalArgumentException("unknow cell vertical alignment "+alignment);
	}

	public static short toRangeApplyBorderType(ApplyBorderType type) {
		switch(type){
		case FULL:
			return BookHelper.BORDER_FULL;
		case EDGE_BOTTOM:
			return BookHelper.BORDER_EDGE_BOTTOM;
		case EDGE_RIGHT:
			return BookHelper.BORDER_EDGE_RIGHT;
		case EDGE_TOP:
			return BookHelper.BORDER_EDGE_TOP;
		case EDGE_LEFT:
			return BookHelper.BORDER_EDGE_LEFT;
		case OUTLINE:
			return BookHelper.BORDER_OUTLINE;
		case INSIDE:
			return BookHelper.BORDER_INSIDE;
		case INSIDE_HORIZONTAL:
			return BookHelper.BORDER_INSIDE_HORIZONTAL;
		case INSIDE_VERTICAL:
			return BookHelper.BORDER_INSIDE_VERTICAL;
		case DIAGONAL:
			return BookHelper.BORDER_DIAGONAL;
		case DIAGONAL_DOWN:
			return BookHelper.BORDER_DIAGONAL_DOWN;
		case DIAGONAL_UP:
			return BookHelper.BORDER_DIAGONAL_UP;
		}
		throw new IllegalArgumentException("unknow cell border apply type "+type);
	}

	public static short toStyleBorderType(BorderType borderType) {
		switch(borderType){
		case NONE:
			return CellStyle.BORDER_NONE;
		case THIN:
			return CellStyle.BORDER_THIN;
		case MEDIUM:
			return CellStyle.BORDER_MEDIUM;
		case DASHED:
			return CellStyle.BORDER_DASHED;
		case HAIR:
			return CellStyle.BORDER_HAIR;
		case THICK:
			return CellStyle.BORDER_THICK;
		case DOUBLE:
			return CellStyle.BORDER_DOUBLE;
		case DOTTED:
			return CellStyle.BORDER_DOTTED;
		case MEDIUM_DASHED:
			return CellStyle.BORDER_MEDIUM_DASHED;
		case DASH_DOT:
			return CellStyle.BORDER_DASH_DOT;
		case MEDIUM_DASH_DOT:
			return CellStyle.BORDER_MEDIUM_DASH_DOT;
		case DASH_DOT_DOT:
			return CellStyle.BORDER_DASH_DOT_DOT;
		case MEDIUM_DASH_DOT_DOT:
			return CellStyle.BORDER_MEDIUM_DASH_DOT_DOT;
		case SLANTED_DASH_DOT:
			return CellStyle.BORDER_SLANTED_DASH_DOT;
		}
		throw new IllegalArgumentException("unknow style border type "+borderType);
	}
	
	public static BorderType toStyleBorderType(short borderType) {
		switch(borderType){
		case CellStyle.BORDER_NONE:
			return BorderType.NONE;
		case CellStyle.BORDER_THIN:
			return BorderType.THIN;
		case CellStyle.BORDER_MEDIUM:
			return BorderType.MEDIUM;
		case CellStyle.BORDER_DASHED:
			return BorderType.DASHED;
		case CellStyle.BORDER_HAIR:
			return BorderType.HAIR;
		case CellStyle.BORDER_THICK:
			return BorderType.THICK;
		case CellStyle.BORDER_DOUBLE:
			return BorderType.DOUBLE;
		case CellStyle.BORDER_DOTTED:
			return BorderType.DOTTED;
		case CellStyle.BORDER_MEDIUM_DASHED:
			return BorderType.MEDIUM_DASHED;
		case CellStyle.BORDER_DASH_DOT:
			return BorderType.DASH_DOT;
		case CellStyle.BORDER_MEDIUM_DASH_DOT:
			return BorderType.MEDIUM_DASH_DOT;
		case CellStyle.BORDER_DASH_DOT_DOT:
			return BorderType.DASH_DOT_DOT;
		case CellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
			return BorderType.MEDIUM_DASH_DOT_DOT;
		case CellStyle.BORDER_SLANTED_DASH_DOT:
			return BorderType.SLANTED_DASH_DOT;
		}
		throw new IllegalArgumentException("unknow style border type "+borderType);
	}
	
	public static BorderStyle toRangeBorderType(BorderType lineStyle) {
		switch(lineStyle){
		case NONE:
			return BorderStyle.NONE;
		case THIN:
			return BorderStyle.THIN;
		case MEDIUM:
			return BorderStyle.MEDIUM;
		case DASHED:
			return BorderStyle.DASHED;
		case HAIR:
			return BorderStyle.HAIR;
		case THICK:
			return BorderStyle.THICK;
		case DOUBLE:
			return BorderStyle.DOUBLE;
		case DOTTED:
			return BorderStyle.DOTTED;
		case MEDIUM_DASHED:
			return BorderStyle.MEDIUM_DASHED;
		case DASH_DOT:
			return BorderStyle.DASH_DOT;
		case MEDIUM_DASH_DOT:
			return BorderStyle.MEDIUM_DASH_DOT;
		case DASH_DOT_DOT:
			return BorderStyle.DASH_DOT_DOT;
		case MEDIUM_DASH_DOT_DOT:
			return BorderStyle.MEDIUM_DASH_DOT_DOT;
		case SLANTED_DASH_DOT:
			return BorderStyle.SLANTED_DASH_DOT;
		}
		throw new IllegalArgumentException("unknow cell border line style "+lineStyle);
	}

	public static int toRangeInsertShift(InsertShift shift) {
		switch(shift){
		case DEFAULT:
			return Range.SHIFT_DEFAULT;
		case DOWN:
			return Range.SHIFT_DOWN;
		case RIGHT:
			return Range.SHIFT_RIGHT;
		}
		throw new IllegalArgumentException("unknow range insert shift "+shift);
	}

	public static int toRangeInsertCopyOrigin(InsertCopyOrigin copyOrigin) {
		switch(copyOrigin){
		case NONE:
			return Range.FORMAT_NONE;
		case LEFT_ABOVE:
			return Range.FORMAT_LEFTABOVE;
		case RIGHT_BELOW:
			return Range.FORMAT_RIGHTBELOW;
		}
		throw new IllegalArgumentException("unknow range insert copy origin "+copyOrigin);
	}
	
	public static int toRangeDeleteShift(DeleteShift shift) {
		switch(shift){
		case DEFAULT:
			return Range.SHIFT_DEFAULT;
		case UP:
			return Range.SHIFT_UP;
		case LEFT:
			return Range.SHIFT_LEFT;
		}
		throw new IllegalArgumentException("unknow range delete shift "+shift);
	}
}
