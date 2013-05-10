package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.common.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.usermodel.charts.ChartGrouping;
import org.zkoss.poi.ss.usermodel.charts.ChartType;
import org.zkoss.zss.api.Range.ApplyBorderType;
import org.zkoss.zss.api.Range.AutoFillType;
import org.zkoss.zss.api.Range.AutoFilterOperation;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Range.PasteOperation;
import org.zkoss.zss.api.Range.PasteType;
import org.zkoss.zss.api.Range.SortDataOption;
import org.zkoss.zss.api.model.CellStyle.Alignment;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.CellStyle.FillPattern;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.VerticalAlignment;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.Chart.Type;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Font.Boldweight;
import org.zkoss.zss.api.model.Font.TypeOffset;
import org.zkoss.zss.api.model.Font.Underline;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.Picture.Format;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.impl.BookHelper;

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
			return XRange.PASTEOP_ADD;
		case PASTEOP_SUB:
			return XRange.PASTEOP_SUB;
		case PASTEOP_MUL:
			return XRange.PASTEOP_MUL;
		case PASTEOP_DIV:
			return XRange.PASTEOP_DIV;
		case PASTEOP_NONE:
			return XRange.PASTEOP_NONE;
		}
		throw new IllegalArgumentException("unknow paste operation "+op);
	}


	public static int toRangePasteTypeNative(PasteType type) {
		assertArgNotNull(type,"paste type");
		switch(type){
		case PASTE_ALL:
			return XRange.PASTE_ALL;
		case PASTE_ALL_EXCEPT_BORDERS:
			return XRange.PASTE_ALL_EXCEPT_BORDERS;
		case PASTE_COLUMN_WIDTHS:
			return XRange.PASTE_COLUMN_WIDTHS;
		case PASTE_COMMENTS:
			return XRange.PASTE_COMMENTS;
		case PASTE_FORMATS:
			return XRange.PASTE_FORMATS;
		case PASTE_FORMULAS:
			return XRange.PASTE_FORMULAS;
		case PASTE_FORMULAS_AND_NUMBER_FORMATS:
			return XRange.PASTE_FORMULAS_AND_NUMBER_FORMATS;
		case PASTE_VALIDATAION:
			return XRange.PASTE_VALIDATAION;
		case PASTE_VALUES:
			return XRange.PASTE_VALUES;
		case PASTE_VALUES_AND_NUMBER_FORMATS:
			return XRange.PASTE_VALUES_AND_NUMBER_FORMATS;
		}
		throw new IllegalArgumentException("unknow paste operation "+type);
	}
	
	public static TypeOffset toFontTypeOffset(short typeOffset){
		switch(typeOffset){
		case org.zkoss.poi.ss.usermodel.Font.SS_NONE:
			return Font.TypeOffset.NONE;
		case org.zkoss.poi.ss.usermodel.Font.SS_SUB:
			return Font.TypeOffset.SUB;
		case org.zkoss.poi.ss.usermodel.Font.SS_SUPER:
			return Font.TypeOffset.SUPER;
		}
		throw new IllegalArgumentException("unknow font type offset "+typeOffset);
	}

	public static short toFontTypeOffset(TypeOffset typeOffset) {
		assertArgNotNull(typeOffset,"typeOffset");
		switch(typeOffset){
		case NONE:
			return org.zkoss.poi.ss.usermodel.Font.SS_NONE;
		case SUB:
			return org.zkoss.poi.ss.usermodel.Font.SS_SUB;
		case SUPER:
			return org.zkoss.poi.ss.usermodel.Font.SS_SUPER;
		}
		throw new IllegalArgumentException("unknow font type offset "+typeOffset);
	}

	public static Underline toFontUnderline(byte underline) {
		switch(underline){
		case org.zkoss.poi.ss.usermodel.Font.U_NONE:
			return Font.Underline.NONE;
		case org.zkoss.poi.ss.usermodel.Font.U_SINGLE:
			return Font.Underline.SINGLE;
		case org.zkoss.poi.ss.usermodel.Font.U_SINGLE_ACCOUNTING:
			return Font.Underline.SINGLE_ACCOUNTING;
		case org.zkoss.poi.ss.usermodel.Font.U_DOUBLE:
			return Font.Underline.DOUBLE;
		case org.zkoss.poi.ss.usermodel.Font.U_DOUBLE_ACCOUNTING:
			return Font.Underline.DOUBLE_ACCOUNTING;
		}
		throw new IllegalArgumentException("unknow font underline "+underline);
	}


	public static byte toFontUnderline(Underline underline) {
		assertArgNotNull(underline,"underline");
		switch(underline){
		case NONE:
			return org.zkoss.poi.ss.usermodel.Font.U_NONE;
		case SINGLE:
			return org.zkoss.poi.ss.usermodel.Font.U_SINGLE;
		case SINGLE_ACCOUNTING:
			return org.zkoss.poi.ss.usermodel.Font.U_SINGLE_ACCOUNTING;
		case DOUBLE:
			return org.zkoss.poi.ss.usermodel.Font.U_DOUBLE;
		case DOUBLE_ACCOUNTING:
			return org.zkoss.poi.ss.usermodel.Font.U_DOUBLE_ACCOUNTING;
		}
		throw new IllegalArgumentException("unknow font underline "+underline);
	}

	public static Boldweight toFontBoldweight(short boldweight) {
		switch(boldweight){
		case org.zkoss.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD:
			return Font.Boldweight.BOLD;
		case org.zkoss.poi.ss.usermodel.Font.BOLDWEIGHT_NORMAL:
			return Font.Boldweight.NORMAL;
		}
		throw new IllegalArgumentException("unknow font boldweight "+boldweight);
	}
	
	public static short toFontBoldweight(Boldweight boldweight) {
		switch(boldweight){
		case BOLD:
			return org.zkoss.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD;
		case NORMAL:
			return org.zkoss.poi.ss.usermodel.Font.BOLDWEIGHT_NORMAL;
		}
		throw new IllegalArgumentException("unknow font boldweight "+boldweight);
	}

	public static FillPattern toStyleFillPattern(short pattern) {
		switch(pattern){
		case org.zkoss.poi.ss.usermodel.CellStyle.NO_FILL:
			return CellStyle.FillPattern.NO_FILL;
		case org.zkoss.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND:
			return CellStyle.FillPattern.SOLID_FOREGROUND;
		case org.zkoss.poi.ss.usermodel.CellStyle.FINE_DOTS:
			return CellStyle.FillPattern.FINE_DOTS;
		case org.zkoss.poi.ss.usermodel.CellStyle.ALT_BARS:
			return CellStyle.FillPattern.ALT_BARS;
		case org.zkoss.poi.ss.usermodel.CellStyle.SPARSE_DOTS:
			return CellStyle.FillPattern.SPARSE_DOTS;
		case org.zkoss.poi.ss.usermodel.CellStyle.THICK_HORZ_BANDS:
			return CellStyle.FillPattern.THICK_HORZ_BANDS;
		case org.zkoss.poi.ss.usermodel.CellStyle.THICK_VERT_BANDS:
			return CellStyle.FillPattern.THICK_VERT_BANDS;
		case org.zkoss.poi.ss.usermodel.CellStyle.THICK_BACKWARD_DIAG:
			return CellStyle.FillPattern.THICK_BACKWARD_DIAG;
		case org.zkoss.poi.ss.usermodel.CellStyle.THICK_FORWARD_DIAG:
			return CellStyle.FillPattern.THICK_FORWARD_DIAG;
		case org.zkoss.poi.ss.usermodel.CellStyle.BIG_SPOTS:
			return CellStyle.FillPattern.BIG_SPOTS;
		case org.zkoss.poi.ss.usermodel.CellStyle.BRICKS:
			return CellStyle.FillPattern.BRICKS;
		case org.zkoss.poi.ss.usermodel.CellStyle.THIN_HORZ_BANDS:
			return CellStyle.FillPattern.THIN_HORZ_BANDS;
		case org.zkoss.poi.ss.usermodel.CellStyle.THIN_VERT_BANDS:
			return CellStyle.FillPattern.THIN_VERT_BANDS;
		case org.zkoss.poi.ss.usermodel.CellStyle.THIN_BACKWARD_DIAG:
			return CellStyle.FillPattern.THIN_BACKWARD_DIAG;
		case org.zkoss.poi.ss.usermodel.CellStyle.THIN_FORWARD_DIAG:
			return CellStyle.FillPattern.THIN_FORWARD_DIAG;
		case org.zkoss.poi.ss.usermodel.CellStyle.SQUARES:
			return CellStyle.FillPattern.SQUARES;
		case org.zkoss.poi.ss.usermodel.CellStyle.DIAMONDS:
			return CellStyle.FillPattern.DIAMONDS;
		case org.zkoss.poi.ss.usermodel.CellStyle.LESS_DOTS:
			return CellStyle.FillPattern.LESS_DOTS;
		case org.zkoss.poi.ss.usermodel.CellStyle.LEAST_DOTS:
			return CellStyle.FillPattern.LEAST_DOTS;
		}
		throw new IllegalArgumentException("unknow pattern type "+pattern);	}
	
	public static short toStyleFillPattern(FillPattern pattern) {
		switch(pattern){
		case NO_FILL:
			return org.zkoss.poi.ss.usermodel.CellStyle.NO_FILL;
		case SOLID_FOREGROUND:
			return org.zkoss.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND;
		case FINE_DOTS:
			return org.zkoss.poi.ss.usermodel.CellStyle.FINE_DOTS;
		case ALT_BARS:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALT_BARS;
		case SPARSE_DOTS:
			return org.zkoss.poi.ss.usermodel.CellStyle.SPARSE_DOTS;
		case THICK_HORZ_BANDS:
			return org.zkoss.poi.ss.usermodel.CellStyle.THICK_HORZ_BANDS;
		case THICK_VERT_BANDS:
			return org.zkoss.poi.ss.usermodel.CellStyle.THICK_VERT_BANDS;
		case THICK_BACKWARD_DIAG:
			return org.zkoss.poi.ss.usermodel.CellStyle.THICK_BACKWARD_DIAG;
		case THICK_FORWARD_DIAG:
			return org.zkoss.poi.ss.usermodel.CellStyle.THICK_FORWARD_DIAG;
		case BIG_SPOTS:
			return org.zkoss.poi.ss.usermodel.CellStyle.BIG_SPOTS;
		case BRICKS:
			return org.zkoss.poi.ss.usermodel.CellStyle.BRICKS;
		case THIN_HORZ_BANDS:
			return org.zkoss.poi.ss.usermodel.CellStyle.THIN_HORZ_BANDS;
		case THIN_VERT_BANDS:
			return org.zkoss.poi.ss.usermodel.CellStyle.THIN_VERT_BANDS;
		case THIN_BACKWARD_DIAG:
			return org.zkoss.poi.ss.usermodel.CellStyle.THIN_BACKWARD_DIAG;
		case THIN_FORWARD_DIAG:
			return org.zkoss.poi.ss.usermodel.CellStyle.THIN_FORWARD_DIAG;
		case SQUARES:
			return org.zkoss.poi.ss.usermodel.CellStyle.SQUARES;
		case DIAMONDS:
			return org.zkoss.poi.ss.usermodel.CellStyle.DIAMONDS;
		case LESS_DOTS:
			return org.zkoss.poi.ss.usermodel.CellStyle.LESS_DOTS;
		case LEAST_DOTS:
			return org.zkoss.poi.ss.usermodel.CellStyle.LEAST_DOTS;
		}
		throw new IllegalArgumentException("unknow pattern type "+pattern);
	}

	public static short toStyleAlignemnt(Alignment alignment) {
		switch(alignment){
		case GENERAL:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_GENERAL;
		case LEFT:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_LEFT;
		case CENTER:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_CENTER;
		case RIGHT:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_RIGHT;
		case FILL:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_FILL;
		case JUSTIFY:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_JUSTIFY;
		case CENTER_SELECTION:
			return org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_CENTER_SELECTION;
		}
		throw new IllegalArgumentException("unknow cell alignment "+alignment);
	}
	public static Alignment toStyleAlignemnt(short alignment) {
		switch(alignment){
		case org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_GENERAL:
			return Alignment.GENERAL;
		case org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_LEFT:
			return Alignment.LEFT;
		case org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_CENTER:
			return Alignment.CENTER;
		case org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_RIGHT:
			return Alignment.RIGHT;
		case org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_FILL:
			return Alignment.FILL;
		case org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_JUSTIFY:
			return Alignment.JUSTIFY;
		case org.zkoss.poi.ss.usermodel.CellStyle.ALIGN_CENTER_SELECTION:
			return Alignment.CENTER_SELECTION;
		}
		throw new IllegalArgumentException("unknow cell alignment "+alignment);
	}
	public static short toStyleVerticalAlignemnt(VerticalAlignment alignment) {
		switch(alignment){
		case TOP:
			return org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_TOP;
		case CENTER:
			return org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_CENTER;
		case BOTTOM:
			return org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_BOTTOM;
		case JUSTIFY:
			return org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_JUSTIFY;
		}
		throw new IllegalArgumentException("unknow cell vertical alignment "+alignment);
	}
	public static VerticalAlignment toStyleVerticalAlignemnt(short alignment) {
		switch(alignment){
		case org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_TOP:
			return VerticalAlignment.TOP;
		case org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_CENTER:
			return VerticalAlignment.CENTER;
		case org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_BOTTOM:
			return VerticalAlignment.BOTTOM;
		case org.zkoss.poi.ss.usermodel.CellStyle.VERTICAL_JUSTIFY:
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
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_NONE;
		case THIN:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_THIN;
		case MEDIUM:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM;
		case DASHED:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DASHED;
		case HAIR:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_HAIR;
		case THICK:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_THICK;
		case DOUBLE:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DOUBLE;
		case DOTTED:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DOTTED;
		case MEDIUM_DASHED:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASHED;
		case DASH_DOT:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DASH_DOT;
		case MEDIUM_DASH_DOT:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASH_DOT;
		case DASH_DOT_DOT:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DASH_DOT_DOT;
		case MEDIUM_DASH_DOT_DOT:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASH_DOT_DOT;
		case SLANTED_DASH_DOT:
			return org.zkoss.poi.ss.usermodel.CellStyle.BORDER_SLANTED_DASH_DOT;
		}
		throw new IllegalArgumentException("unknow style border type "+borderType);
	}
	
	public static BorderType toStyleBorderType(short borderType) {
		switch(borderType){
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_NONE:
			return BorderType.NONE;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_THIN:
			return BorderType.THIN;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM:
			return BorderType.MEDIUM;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DASHED:
			return BorderType.DASHED;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_HAIR:
			return BorderType.HAIR;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_THICK:
			return BorderType.THICK;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DOUBLE:
			return BorderType.DOUBLE;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DOTTED:
			return BorderType.DOTTED;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASHED:
			return BorderType.MEDIUM_DASHED;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DASH_DOT:
			return BorderType.DASH_DOT;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASH_DOT:
			return BorderType.MEDIUM_DASH_DOT;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_DASH_DOT_DOT:
			return BorderType.DASH_DOT_DOT;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
			return BorderType.MEDIUM_DASH_DOT_DOT;
		case org.zkoss.poi.ss.usermodel.CellStyle.BORDER_SLANTED_DASH_DOT:
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
			return XRange.SHIFT_DEFAULT;
		case DOWN:
			return XRange.SHIFT_DOWN;
		case RIGHT:
			return XRange.SHIFT_RIGHT;
		}
		throw new IllegalArgumentException("unknow range insert shift "+shift);
	}

	public static int toRangeInsertCopyOrigin(InsertCopyOrigin copyOrigin) {
		switch(copyOrigin){
		case NONE:
			return XRange.FORMAT_NONE;
		case LEFT_ABOVE:
			return XRange.FORMAT_LEFTABOVE;
		case RIGHT_BELOW:
			return XRange.FORMAT_RIGHTBELOW;
		}
		throw new IllegalArgumentException("unknow range insert copy origin "+copyOrigin);
	}
	
	public static int toRangeDeleteShift(DeleteShift shift) {
		switch(shift){
		case DEFAULT:
			return XRange.SHIFT_DEFAULT;
		case UP:
			return XRange.SHIFT_UP;
		case LEFT:
			return XRange.SHIFT_LEFT;
		}
		throw new IllegalArgumentException("unknow range delete shift "+shift);
	}

	public static int toRangeSortDataOption(SortDataOption dataOption) {
		switch(dataOption){
		case TEXT_AS_NUMBERS:
			return BookHelper.SORT_TEXT_AS_NUMBERS;
		}
		throw new IllegalArgumentException("unknow sort data option "+dataOption);
	}

	public static int toRangeAutoFilterOperation(AutoFilterOperation filterOp) {
		switch(filterOp){
		case AND:
			return AutoFilter.FILTEROP_AND;
		case OR:
			return AutoFilter.FILTEROP_OR;
		case TOP10:
			return AutoFilter.FILTEROP_TOP10;
		case TOP10PERCENT:
			return AutoFilter.FILTEROP_TOP10PERCENT;
		case BOTTOM10:
			return AutoFilter.FILTEROP_BOTTOM10;
		case BOTOOM10PERCENT:
			return AutoFilter.FILTEROP_BOTOOM10PERCENT;
		case VALUES:
			return AutoFilter.FILTEROP_VALUES;
		}
		throw new IllegalArgumentException("unknow autofilter operation "+filterOp);
	}

	public static int toRangeAutoFillType(AutoFillType fillType) {
		switch(fillType){
		case COPY:
			return XRange.FILL_COPY;
		case DAYS:
			return XRange.FILL_DAYS;
		case DEFAULT:
			return XRange.FILL_DEFAULT;
		case FORMATS:
			return XRange.FILL_FORMATS;
		case MONTHS:
			return XRange.FILL_MONTHS;
		case SERIES:
			return XRange.FILL_SERIES;
		case VALUES:
			return XRange.FILL_VALUES;
		case WEEKDAYS:
			return XRange.FILL_WEEKDAYS;
		case YEARS:
			return XRange.FILL_YEARS;
		case GROWTH_TREND:
			return XRange.FILL_GROWTH_TREND;
		case LINER_TREND:
			return XRange.FILL_LINER_TREND;
		}
		throw new IllegalArgumentException("unknow autofill type "+fillType);
	}

	public static int toHyperlinkType(HyperlinkType type) {
		switch(type){
		case URL:
			return Hyperlink.LINK_URL;
		case DOCUMENT:
			return Hyperlink.LINK_DOCUMENT;
		case EMAIL:
			return Hyperlink.LINK_EMAIL;
		case FILE:
			return Hyperlink.LINK_FILE;
		}
		throw new IllegalArgumentException("unknow hyperlink type "+type);
	}
	public static HyperlinkType toHyperlinkType(int type) {
		switch(type){
		case Hyperlink.LINK_URL:
			return HyperlinkType.URL;
		case Hyperlink.LINK_DOCUMENT:
			return HyperlinkType.DOCUMENT;
		case Hyperlink.LINK_EMAIL:
			return HyperlinkType.EMAIL;
		case Hyperlink.LINK_FILE:
			return HyperlinkType.FILE;
		}
		throw new IllegalArgumentException("unknow hyperlink type "+type);
	}

	public static int toPictureFormat(Format format) {
		switch(format){
		case EMF:
			return Workbook.PICTURE_TYPE_EMF;
		case WMF:
			return Workbook.PICTURE_TYPE_WMF;
		case PICT:
			return Workbook.PICTURE_TYPE_PICT;
		case JPEG:
			return Workbook.PICTURE_TYPE_JPEG;
		case PNG:
			return Workbook.PICTURE_TYPE_PNG;
		case DIB:
			return Workbook.PICTURE_TYPE_DIB;
		}
		throw new IllegalArgumentException("unknow pciture format "+format);
	}

	public static ChartType toChartType(Type type) {
		switch(type){
		case Area3D:
			return ChartType.Area3D;
		case Area:
			return ChartType.Area3D;
		case Bar3D:
			return ChartType.Bar3D;
		case Bar:
			return ChartType.Bar;
		case Bubble:
			return ChartType.Bubble;
		case Column:
			return ChartType.Column;
		case Column3D:
			return ChartType.Column3D;
		case Doughnut:
			return ChartType.Doughnut;
		case Line3D:
			return ChartType.Line3D;
		case Line:
			return ChartType.Line;
		case OfPie:
			return ChartType.OfPie;
		case Pie3D:
			return ChartType.Pie3D;
		case Pie:
			return ChartType.Pie;
		case Radar:
			return ChartType.Radar;
		case Scatter:
			return ChartType.Scatter;
		case Stock:
			return ChartType.Stock;
		case Surface3D:
			return ChartType.Surface3D;
		case Surface:
			return ChartType.Surface;
		}
		throw new IllegalArgumentException("unknow chart type "+type);
	}

	public static ChartGrouping toChartGrouping(Grouping grouping) {
		switch(grouping){
		case STANDARD:
			return ChartGrouping.STANDARD;
		case STACKED:
			return ChartGrouping.STACKED;
		case PERCENT_STACKED:
			return ChartGrouping.PERCENT_STACKED;
		case CLUSTERED:
			return ChartGrouping.CLUSTERED;//bar only
		}
		throw new IllegalArgumentException("unknow grouping "+grouping);
	}

	public static org.zkoss.poi.ss.usermodel.charts.LegendPosition toLegendPosition(
			LegendPosition pos) {
		switch(pos){
		case BOTTOM:
			return org.zkoss.poi.ss.usermodel.charts.LegendPosition.BOTTOM;
		case LEFT:
			return org.zkoss.poi.ss.usermodel.charts.LegendPosition.LEFT;
		case RIGHT:
			return org.zkoss.poi.ss.usermodel.charts.LegendPosition.RIGHT;
		case TOP:
			return org.zkoss.poi.ss.usermodel.charts.LegendPosition.TOP;
		case TOP_RIGHT:
			return org.zkoss.poi.ss.usermodel.charts.LegendPosition.TOP_RIGHT;
		}
		throw new IllegalArgumentException("unknow legend position "+pos);
	}
}
