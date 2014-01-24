/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2014/1/24 , Created by Kuro
}}IS_NOTE

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/

package org.zkoss.zss.ngapi.impl.imexp;

import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.PrintSetup;
import org.zkoss.zss.ngmodel.NPrintSetup;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.Boldweight;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.NFont.Underline;

/**
 * enum utility for exporter ZSS model to POI model
 * 
 * @author kuro
 * @since 3.5.0
 */
public class ExporterEnumUtil {
	
	public static int toPoiFilterOperator(FilterOp operator){
		switch(operator){
			case AND:
				return AutoFilter.FILTEROP_AND;
			case BOTTOM10:
				return AutoFilter.FILTEROP_BOTTOM10;
			case BOTOOM10_PERCENT:
				return AutoFilter.FILTEROP_BOTOOM10PERCENT;
			case OR:
				return AutoFilter.FILTEROP_OR;
			case TOP10:
				return AutoFilter.FILTEROP_TOP10;
			case TOP10_PERCENT:
				return AutoFilter.FILTEROP_TOP10PERCENT;
			case VALUES:
			default:
				return AutoFilter.FILTEROP_VALUES;
		}
	}

	public static short toPoiPaperSize(NPrintSetup.PaperSize paperSize) {
		switch(paperSize) {
		case A3:
			return PrintSetup.A3_PAPERSIZE;
		case A4_EXTRA:
			return PrintSetup.A4_EXTRA_PAPERSIZE;
		case A4_PLUS:
			return PrintSetup.A4_PLUS_PAPERSIZE;
		case A4_ROTATED:
			return PrintSetup.A4_ROTATED_PAPERSIZE;
		case A4_SMALL:
			return PrintSetup.A4_SMALL_PAPERSIZE;
		case A4:
			return PrintSetup.A4_PAPERSIZE;
		case A4_TRANSVERSE:
			return PrintSetup.A4_TRANSVERSE_PAPERSIZE;
		case A5:
			return PrintSetup.A5_PAPERSIZE;
		case B4:
			return PrintSetup.B4_PAPERSIZE;
		case B5:
			return PrintSetup.B5_PAPERSIZE;
		case ELEVEN_BY_SEVENTEEN:
			return PrintSetup.ELEVEN_BY_SEVENTEEN_PAPERSIZE;
		case ENVELOPE_10:
			return PrintSetup.ENVELOPE_10_PAPERSIZE;
		case ENVELOPE_9:
			return PrintSetup.ENVELOPE_9_PAPERSIZE;
		case ENVELOPE_C3:
			return PrintSetup.ENVELOPE_C3_PAPERSIZE;
		case ENVELOPE_C4:
			return PrintSetup.ENVELOPE_C4_PAPERSIZE;
		case ENVELOPE_C5:
			return PrintSetup.ENVELOPE_C5_PAPERSIZE;
		case ENVELOPE_C6:
			return PrintSetup.ENVELOPE_C6_PAPERSIZE;
		case ENVELOPE_CS:
			return PrintSetup.ENVELOPE_CS_PAPERSIZE;
		case ENVELOPE_DL:
			return PrintSetup.ENVELOPE_DL_PAPERSIZE;
		case ENVELOPE_MONARCH:
			return PrintSetup.ENVELOPE_MONARCH_PAPERSIZE;
		case EXECUTIVE:
			return PrintSetup.EXECUTIVE_PAPERSIZE;
		case FOLIO8:
			return PrintSetup.FOLIO8_PAPERSIZE;
		case LEDGER:
			return PrintSetup.LEDGER_PAPERSIZE;
		case LEGAL:
			return PrintSetup.LEGAL_PAPERSIZE;
		case LETTER:
			return PrintSetup.LETTER_PAPERSIZE;
		case LETTER_ROTATED:
			return PrintSetup.LETTER_ROTATED_PAPERSIZE;
		case LETTER_SMALL:
			return PrintSetup.LETTER_SMALL_PAGESIZE;
		case NOTE8:
			return PrintSetup.NOTE8_PAPERSIZE;
		case QUARTO:
			return PrintSetup.QUARTO_PAPERSIZE;
		case STATEMENT:
			return PrintSetup.STATEMENT_PAPERSIZE;
		case TABLOID:
			return PrintSetup.TABLOID_PAPERSIZE;
		case TEN_BY_FOURTEEN:
			return PrintSetup.TEN_BY_FOURTEEN_PAPERSIZE;
			default:
				return PrintSetup.A4_PAPERSIZE;
		}
	}
	
	public static short toPoiVerticalAlignment(VerticalAlignment vAlignment) {
		switch(vAlignment) {
			case BOTTOM:
				return CellStyle.VERTICAL_BOTTOM;
			case CENTER:
				return CellStyle.VERTICAL_CENTER;
			case JUSTIFY:
				return CellStyle.VERTICAL_JUSTIFY;
			case TOP:
			default:
				return CellStyle.VERTICAL_TOP;
		}
	}
	
	public static short toPoiFillPattern(FillPattern fillPattern) {
		switch(fillPattern) {
		case ALT_BARS:
			return CellStyle.ALT_BARS;
		case BIG_SPOTS:
			return CellStyle.BIG_SPOTS;
		case BRICKS:
			return CellStyle.BRICKS;
		case DIAMONDS:
			return CellStyle.DIAMONDS;
		case FINE_DOTS:
			return CellStyle.FINE_DOTS;
		case LEAST_DOTS:
			return CellStyle.LEAST_DOTS;
		case LESS_DOTS:
			return CellStyle.LESS_DOTS;
		case SOLID_FOREGROUND:
			return CellStyle.SOLID_FOREGROUND;
		case SPARSE_DOTS:
			return CellStyle.SPARSE_DOTS;
		case SQUARES:
			return CellStyle.SQUARES;
		case THICK_BACKWARD_DIAG:
			return CellStyle.THICK_BACKWARD_DIAG;
		case THICK_FORWARD_DIAG:
			return CellStyle.THICK_FORWARD_DIAG;
		case THICK_HORZ_BANDS:
			return CellStyle.THICK_HORZ_BANDS;
		case THICK_VERT_BANDS:
			return CellStyle.THICK_VERT_BANDS;
		case THIN_BACKWARD_DIAG:
			return CellStyle.THIN_BACKWARD_DIAG;
		case THIN_FORWARD_DIAG:
			return CellStyle.THIN_FORWARD_DIAG;
		case THIN_HORZ_BANDS:
			return CellStyle.THIN_HORZ_BANDS;
		case THIN_VERT_BANDS:
			return CellStyle.THIN_VERT_BANDS;
		case NO_FILL:
		default:
			return CellStyle.NO_FILL;
		}
	}
	
	public static short toPoiBorderType(BorderType borderType) {
		switch(borderType) {
		case DASH_DOT:
			return CellStyle.BORDER_DASH_DOT;
		case DASHED:
			return CellStyle.BORDER_DASHED;
		case DOTTED:
			return CellStyle.BORDER_DOTTED;
		case DOUBLE:
			return CellStyle.BORDER_DOUBLE;
		case HAIR:
			return CellStyle.BORDER_HAIR;
		case MEDIUM:
			return CellStyle.BORDER_MEDIUM;
		case MEDIUM_DASH_DOT:
			return CellStyle.BORDER_DASH_DOT;
		case MEDIUM_DASH_DOT_DOT:
			return CellStyle.BORDER_DASH_DOT_DOT;
		case MEDIUM_DASHED:
			return CellStyle.BORDER_MEDIUM_DASHED;
		case SLANTED_DASH_DOT:
			return CellStyle.BORDER_SLANTED_DASH_DOT;
		case THICK:
			return CellStyle.BORDER_THICK;
		case THIN:
			return CellStyle.BORDER_THIN;
		case DASH_DOT_DOT:
			return CellStyle.BORDER_DASH_DOT_DOT;
		case NONE:
		default:
			return CellStyle.BORDER_NONE;
		}
	}
	
	public static short toPoiAlignment(Alignment alignment) {
		switch(alignment) {
		case CENTER:
			return CellStyle.ALIGN_CENTER;
		case FILL:
			return CellStyle.ALIGN_FILL;
		case JUSTIFY:
			return CellStyle.ALIGN_JUSTIFY;
		case RIGHT:
			return CellStyle.ALIGN_RIGHT;
		case LEFT:
			return CellStyle.ALIGN_LEFT;
		case CENTER_SELECTION:
			return CellStyle.ALIGN_CENTER_SELECTION;
		case GENERAL:
			default:
			return CellStyle.ALIGN_GENERAL;
		}
	}
	
	public static short toPoiBoldweight(Boldweight bold) {
		switch(bold) {
			case BOLD:
				return Font.BOLDWEIGHT_BOLD;
			case NORMAL:
			default:
				return Font.BOLDWEIGHT_NORMAL;
		}
	}
	
	public static short toPoiTypeOffset(TypeOffset typeOffset) {
		switch(typeOffset) {
			case SUPER:
				return Font.SS_SUPER;
			case SUB:
				return Font.SS_SUB;
			case NONE:
			default:
				return Font.SS_NONE;
		}
	}
	
	public static byte toPoiUnderline(Underline underline) {
		switch(underline) {
			case SINGLE:
				return Font.U_SINGLE;
			case DOUBLE:
				return Font.U_DOUBLE;
			case DOUBLE_ACCOUNTING:
				return Font.U_DOUBLE_ACCOUNTING;
			case SINGLE_ACCOUNTING:
				return Font.U_SINGLE_ACCOUNTING;
			case NONE:
			default:
				return Font.U_NONE;
		}
	}
	
}
