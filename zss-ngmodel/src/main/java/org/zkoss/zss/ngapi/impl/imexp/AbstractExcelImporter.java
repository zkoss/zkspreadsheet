/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/13 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NCellStyle.*;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.*;


/**
 * Contains common importing behavior for both XLSX and XLS.
 * @author Hawk
 *
 */
abstract public class AbstractExcelImporter extends AbstractImporter {
	/*
	 * reference BookHelper.getFontCSSStyle()
	 */
	protected Underline convertUnderline(byte underline){
		switch(underline){
		case Font.U_SINGLE:
			return NFont.Underline.SINGLE;
		case Font.U_DOUBLE:
			return NFont.Underline.DOUBLE;
		case Font.U_SINGLE_ACCOUNTING:
			return NFont.Underline.SINGLE_ACCOUNTING;
		case Font.U_DOUBLE_ACCOUNTING:
			return NFont.Underline.DOUBLE_ACCOUNTING;
		case Font.U_NONE:
		default:
			return NFont.Underline.NONE;
		}
	}
	
	protected TypeOffset convertTypeOffset(short typeOffset){
		switch(typeOffset){
		case Font.SS_SUB:
			return TypeOffset.SUB;
		case Font.SS_SUPER:
			return TypeOffset.SUPER;
		case Font.SS_NONE:
		default:
			return TypeOffset.NONE;
		}
	}
	
	protected BorderType convertBorderType(short poiBorder) {
		switch (poiBorder) {
			case CellStyle.BORDER_DASH_DOT:
				return BorderType.DASH_DOT;
			case CellStyle.BORDER_DASH_DOT_DOT:
				return BorderType.DASH_DOT_DOT;
			case CellStyle.BORDER_DASHED:
				return BorderType.DASHED;
			case CellStyle.BORDER_DOTTED:
				return BorderType.DOTTED;
			case CellStyle.BORDER_NONE:
				return BorderType.NONE;
			case CellStyle.BORDER_THIN:
			default:
				return BorderType.THIN; //unsupported border type
		}				
	}
	
	protected Alignment convertAlignment(short poiAlignment) {
		switch (poiAlignment) {
			case CellStyle.ALIGN_LEFT:
				return Alignment.LEFT;
			case CellStyle.ALIGN_RIGHT:
				return Alignment.RIGHT;
			case CellStyle.ALIGN_CENTER:
			case CellStyle.ALIGN_CENTER_SELECTION:
				return Alignment.CENTER;
			case CellStyle.ALIGN_FILL:
				return Alignment.FILL;
			case CellStyle.ALIGN_JUSTIFY:
				return Alignment.JUSTIFY;
			case CellStyle.ALIGN_GENERAL:
			default:	
				return Alignment.GENERAL;
		}
	}
	
	protected VerticalAlignment convertVerticalAlignment(short vAlignment) {
		switch (vAlignment) {
			case CellStyle.VERTICAL_TOP:
				return VerticalAlignment.TOP;
			case CellStyle.VERTICAL_CENTER:
				return VerticalAlignment.CENTER;
			case CellStyle.VERTICAL_JUSTIFY:
				return VerticalAlignment.JUSTIFY;
			case CellStyle.VERTICAL_BOTTOM:
			default:	
				return VerticalAlignment.BOTTOM;
		}
	}

	//TODO more error codes
	protected ErrorValue convertErrorCode(byte errorCellValue) {
		switch (errorCellValue){
			case ErrorConstants.ERROR_NAME:
				return new ErrorValue(ErrorValue.INVALID_NAME);
			case ErrorConstants.ERROR_VALUE:
				return new ErrorValue(ErrorValue.INVALID_VALUE);
			default:
				//TODO log it
				return new ErrorValue(ErrorValue.INVALID_NAME);
		}
		
	}
	
	protected FillPattern convertPoiFillPattern(short poiFillPattern){
		switch(poiFillPattern){
			case CellStyle.SOLID_FOREGROUND:
				return FillPattern.SOLID_FOREGROUND;
			case CellStyle.FINE_DOTS:
				return FillPattern.FINE_DOTS;
			case CellStyle.ALT_BARS:
				return FillPattern.ALT_BARS;
			case CellStyle.SPARSE_DOTS:
				return FillPattern.SPARSE_DOTS;
			case CellStyle.THICK_HORZ_BANDS:
				return FillPattern.THICK_HORZ_BANDS;
			case CellStyle.THICK_VERT_BANDS:
				return FillPattern.THICK_VERT_BANDS;
			case CellStyle.THICK_BACKWARD_DIAG:
				return FillPattern.THICK_BACKWARD_DIAG;
			case CellStyle.THICK_FORWARD_DIAG:
				return FillPattern.THICK_FORWARD_DIAG;
			case CellStyle.BIG_SPOTS:
				return FillPattern.BIG_SPOTS;
			case CellStyle.BRICKS:
				return FillPattern.BRICKS;
			case CellStyle.THIN_HORZ_BANDS:
				return FillPattern.THIN_HORZ_BANDS;
			case CellStyle.THIN_VERT_BANDS:
				return FillPattern.THIN_VERT_BANDS;
			case CellStyle.THIN_BACKWARD_DIAG:
				return FillPattern.THIN_BACKWARD_DIAG;
			case CellStyle.THIN_FORWARD_DIAG:
				return FillPattern.THIN_FORWARD_DIAG;
			case CellStyle.SQUARES:
				return FillPattern.SQUARES;
			case CellStyle.DIAMONDS:
				return FillPattern.DIAMONDS;
			case CellStyle.LESS_DOTS:
				return FillPattern.LESS_DOTS;
			case CellStyle.LEAST_DOTS:
				return FillPattern.LEAST_DOTS;
			case CellStyle.NO_FILL:
			default:
			return FillPattern.NO_FILL;
		}
	}
}
