/* BookHelper.java

	Purpose:
		
	Description:
		
	History:
		Mar 24, 2010 5:42:58 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.poi.hssf.record.CellValueRecordInterface;
import org.zkoss.poi.hssf.record.FormulaRecord;
import org.zkoss.poi.hssf.record.aggregates.FormulaRecordAggregate;
import org.zkoss.poi.hssf.record.formula.Area3DPtg;
import org.zkoss.poi.hssf.record.formula.AreaPtgBase;
import org.zkoss.poi.hssf.record.formula.Ptg;
import org.zkoss.poi.hssf.record.formula.RefPtgBase;
import org.zkoss.poi.hssf.usermodel.HSSFCell;
import org.zkoss.poi.hssf.usermodel.HSSFCellHelper;
import org.zkoss.poi.hssf.usermodel.HSSFCellStyle;
import org.zkoss.poi.hssf.usermodel.HSSFDataFormat;
import org.zkoss.poi.hssf.usermodel.HSSFDataValidation;
import org.zkoss.poi.hssf.usermodel.HSSFFont;
import org.zkoss.poi.hssf.usermodel.HSSFPalette;
import org.zkoss.poi.hssf.usermodel.HSSFRichTextString;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.hssf.util.HSSFColor;
import org.zkoss.poi.hssf.util.PaneInformation;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.format.CellFormat;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaParsingWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.PtgShifter;
import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Comment;
import org.zkoss.poi.ss.usermodel.CreationHelper;
import org.zkoss.poi.ss.usermodel.DataFormatter;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.DataValidationConstraint;
import org.zkoss.poi.ss.usermodel.DataValidationHelper;
import org.zkoss.poi.ss.usermodel.Drawing;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.ss.util.CellRangeAddressList;
import org.zkoss.poi.ss.util.NumberToTextConverter;
import org.zkoss.poi.xssf.model.ThemesTable;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.poi.xssf.usermodel.XSSFCellStyle;
import org.zkoss.poi.xssf.usermodel.XSSFColor;
import org.zkoss.poi.xssf.usermodel.XSSFDataFormat;
import org.zkoss.poi.xssf.usermodel.XSSFDataValidation;
import org.zkoss.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.zkoss.poi.xssf.usermodel.XSSFFont;
import org.zkoss.poi.xssf.usermodel.XSSFRichTextString;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelContext;
import org.zkoss.xel.util.SimpleXelContext;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.RefSheet;
import org.zkoss.zss.engine.event.SSDataEvent;
import org.zkoss.zss.engine.impl.AreaRefImpl;
import org.zkoss.zss.engine.impl.ChangeInfo;
import org.zkoss.zss.engine.impl.MergeChange;
import org.zkoss.zss.engine.impl.RefSheetImpl;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.BookSeries;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.ui.impl.Styles;

/**
 * Internal User only. Helper class to handle the two version of Books.
 * @author henrichen
 */
public final class BookHelper {
	public static final String AUTO_COLOR = "AUTO_COLOR";

	//@see #setBorders()
	public static final short BORDER_EDGE_BOTTOM		= 0x01;
	public static final short BORDER_EDGE_RIGHT			= 0x02;
	public static final short BORDER_EDGE_TOP			= 0x04;
	public static final short BORDER_EDGE_LEFT			= 0x08;
	public static final short BORDER_INSIDE_HORIZONTAL	= 0x10;
	public static final short BORDER_INSIDE_VERTICAL 	= 0x20;
	public static final short BORDER_DIAGONAL_DOWN		= 0x40;
	public static final short BORDER_DIAGONAL_UP		= 0x80;
	
	public static final short BORDER_FULL		= 0xFF;
	public static final short BORDER_OUTLINE	= 0x0F;
	public static final short BORDER_INSIDE		= 0x30;
	public static final short BORDER_DIAGONAL	= 0xC0;
	
	//@see #sort()
	public static final int SORT_NORMAL_DEFAULT = 0;
	public static final int SORT_TEXT_AS_NUMBERS = 1;
	public static final int SORT_HEADER_NO  = 0;
	public static final int SORT_HEADER_YES = 1;	
	
	//inner pasteType for #paste & #fill
	public final static int INNERPASTE_NUMBER_FORMATS = 0x01;
	public final static int INNERPASTE_BORDERS = 0x02;
	public final static int INNERPASTE_OTHER_FORMATS = 0x04;
	public final static int INNERPASTE_VALUES = 0x08;
	public final static int INNERPASTE_FORMULAS = 0x10;
	public final static int INNERPASTE_VALUES_AND_FORMULAS = INNERPASTE_FORMULAS + INNERPASTE_VALUES;
	public final static int INNERPASTE_COMMENTS = 0x20;
	public final static int INNERPASTE_VALIDATION = 0x40;
	public final static int INNERPASTE_COLUMN_WIDTHS = 0x80;
	public final static int INNERPASTE_FORMATS = INNERPASTE_NUMBER_FORMATS + INNERPASTE_BORDERS + INNERPASTE_OTHER_FORMATS;
	
	private final static int INNERPASTE_FILL_COPY = INNERPASTE_FORMATS + INNERPASTE_VALUES_AND_FORMULAS + INNERPASTE_VALIDATION;   
	private final static int INNERPASTE_FILL_VALUE = INNERPASTE_VALUES_AND_FORMULAS + INNERPASTE_VALIDATION;   
	private final static int INNERPASTE_FILL_FORMATS = INNERPASTE_FORMATS;   

	//inner pasteOp for #paste & #fill
	public final static int PASTEOP_ADD = 1;
	public final static int PASTEOP_SUB = 2;
	public final static int PASTEOP_MUL = 3;
	public final static int PASTEOP_DIV = 4;
	public final static int PASTEOP_NONE = 0;

	//inner fillType for #fill
	public final static int FILL_DEFAULT = 0x01; //system determine
	public final static int FILL_FORMATS = 0x02; //formats only
	public final static int FILL_VALUES = 0x04; //value+formula+validation+hyperlink (no comment)
	public final static int FILL_COPY = 0x06; //value+formula+validation+hyperlink, formats
	public final static int FILL_DAYS = 0x10;
	public final static int FILL_WEEKDAYS = 0x20;
	public final static int FILL_MONTHS = 0x30;
	public final static int FILL_YEARS = 0x40;
	public final static int FILL_GROWTH_TREND = 0x100; //multiplicative relation
	public final static int FILL_LINER_TREND = 0x200; //additive relation
	public final static int FILL_SERIES = FILL_LINER_TREND;
	
	//inner filterOp for #filter
	public final static int FILTEROP_AND = 0x01;
	public final static int FILTEROP_BOTTOM10 = 0x02;
	public final static int FILTEROP_BOTOOM10PERCENT = 0x03;
	public final static int FILTEROP_OR = 0x04;
	public final static int FILTEROP_TOP10 = 0x05;
	public final static int FILTEROP_TOP10PERCENT = 0x06;
	
	public static RefBook getRefBook(Book book) {
		return book instanceof HSSFBookImpl ? 
			((HSSFBookImpl)book).getOrCreateRefBook():
			((XSSFBookImpl)book).getOrCreateRefBook();
	}
	
	public static RefBook getOrCreateRefBook(Book book) {
		return book instanceof HSSFBookImpl ? 
			((HSSFBookImpl)book).getOrCreateRefBook():
			((XSSFBookImpl)book).getOrCreateRefBook();
	}
	
	public static RefSheet getRefSheet(Book book, Sheet sheet) {
		final RefBook refBook = getRefBook(book);
		return refBook.getOrCreateRefSheet(sheet.getSheetName());
	}
	
	public static BookSeries getBooks(Book book) {
		return book instanceof HSSFBookImpl ? 
				((HSSFBookImpl)book).getBookSeries():
				((XSSFBookImpl)book).getBookSeries();
	}
	
	public static void setBooks(Book book, BookSeries books) {
		if (book instanceof HSSFBookImpl) 
			((HSSFBookImpl)book).setBookSeries(books);
		else
			((XSSFBookImpl)book).setBookSeries(books);
	}
	
	public static Book getBook(Book book, String bookname) {
		final BookSeries books = BookHelper.getBooks(book);
		return bookname == null || book.getBookName().equals(bookname) ? book : books != null ? books.getBook(bookname) : null;
	}
	
	public static Book getBook(Book book, RefSheet refSheet) {
		final String bookName = refSheet.getOwnerBook().getBookName();
		return BookHelper.getBook(book, bookName);
	}
	
	public static Book getBook(Sheet sheet, RefSheet refSheet) {
		return BookHelper.getBook((Book)sheet.getWorkbook(), refSheet);
	}
	
	public static Sheet getSheet(Sheet sheet, RefSheet refSheet) {
		final Book targetBook = BookHelper.getBook(sheet, refSheet);
		final String sheetName = refSheet.getSheetName();
		return targetBook.getSheet(sheetName);
	}
	
	public static VariableResolver getVariableResolver(Book book) {
		if (book instanceof HSSFBookImpl) 
			return ((HSSFBookImpl)book).getVariableResolver();
		else
			return ((XSSFBookImpl)book).getVariableResolver();
	}
	
	public static FunctionMapper getFunctionMapper(Book book) {
		if (book instanceof HSSFBookImpl) 
			return ((HSSFBookImpl)book).getFunctionMapper();
		else
			return ((XSSFBookImpl)book).getFunctionMapper();
	}
	
	/*package*/ static void clearFormulaCache(Book book, Set<Ref> all) {
		if (all != null) {
			for(Ref ref : all) {
				final RefSheet refSheet = ref.getOwnerSheet();
				final Book bookTarget = BookHelper.getBook(book, refSheet);
				final Cell cell = getCell(book, ref.getTopRow(), ref.getLeftCol(), refSheet);
				if (cell != null)
					bookTarget.getFormulaEvaluator().notifySetFormula(cell); 
			}
		}
	}
	
	/*package*/ static void reevaluate(Book book, Set<Ref> last) {
		if (last != null) {
			for(Ref ref : last) {
				final RefSheet refSheet = ref.getOwnerSheet();
				//locate the model book and sheet of the refSheet
				final Book bookTarget = BookHelper.getBook(book, refSheet);
				final Cell cell = getCell(book, ref.getTopRow(), ref.getLeftCol(), refSheet);
				if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
					evaluate(bookTarget, cell);
				}
			}
		}
	}
	
	public static void notifyCellChanges(Book book, Set<Ref> all) {
		if (all != null) {
			for(Ref ref : all) {
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				refBook.publish(new SSDataEvent(SSDataEvent.ON_CONTENTS_CHANGE, ref, SSDataEvent.MOVE_NO));
			}
		}
	}
	
	public static void notifySizeChanges(Book book, Set<Ref> all) {
		if (all != null) {
			for(Ref ref : all) {
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				refBook.publish(new SSDataEvent(SSDataEvent.ON_SIZE_CHANGE, ref, SSDataEvent.MOVE_NO));
			}
		}
	}
	
	/*package*/ static void notifyGridlines(Book book, Set<Ref> all, boolean show) {
		if (all != null) {
			for(Ref ref : all) {
				final RefSheet refSheet = ref.getOwnerSheet();
				final RefBook refBook = refSheet.getOwnerBook();
				refBook.publish(new SSDataEvent(SSDataEvent.ON_DISPLAY_GRIDLINES, ref, show));
			}
		}
	}
	
	public static void reevaluateAndNotify(Book book, Set<Ref> last, Set<Ref> all) {
		//clear cached formula value
		clearFormulaCache(book, all);
		
		//re-evaluate all required formula cells
		reevaluate(book, last);

		//notify all changed cells
		notifyCellChanges(book, all);
	}

	public static CellValue evaluate(Book book, Cell cell) {
		final XelContext old = XelContextHolder.getXelContext();
		try {
			final VariableResolver resolver = BookHelper.getVariableResolver(book);
			final FunctionMapper mapper = BookHelper.getFunctionMapper(book);
			final XelContext ctx = new SimpleXelContext(resolver, mapper);
			ctx.setAttribute("zkoss.zss.CellType", Object.class);
			XelContextHolder.setXelContext(ctx);
			final CellValue cv = book.getFormulaEvaluator().evaluate(cell);
			//set back into Cell formula record(update value and cachedFormulaResultType)
			setCellValue(cell, cv);
			return cv;
		} finally {
			XelContextHolder.setXelContext(old);
		}
	}
	
	private static void setCellValue(Cell cell, CellValue cv) {
		final int type = cv.getCellType();
		switch(type) {
		case Cell.CELL_TYPE_BLANK:
			cell.setCellValue((String)null);
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			cell.setCellValue(cv.getBooleanValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			cell.setCellErrorValue(cv.getErrorValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			cell.setCellValue(cv.getNumberValue());
			break;
		case Cell.CELL_TYPE_STRING:
			cell.setCellValue(cv.getStringValue());
			break;
		}
		Hyperlink hyperlink = cv.getHyperlink();
		if (hyperlink != null) {
			if (cell instanceof HSSFCell)
				((HSSFCell)cell).setEvalHyperlink(hyperlink);
			else
				((XSSFCell)cell).setEvalHyperlink(hyperlink);
		}
	}
	
	private static Cell getCell(Book book, int rowIndex, int colIndex, RefSheet refSheet) {
		//locate the model book and sheet of the refSheet
		final Book bookTarget = BookHelper.getBook(book, refSheet);
		if (bookTarget != null) {
			final Sheet sheet0 = bookTarget.getSheet(refSheet.getSheetName());
			if (sheet0 != null)
				return getCell(sheet0, rowIndex, colIndex);
		}
		return null;
	}
	
	public static Cell getCell(Sheet sheet, int rowIndex, int colIndex) {
		final Row row = sheet.getRow(rowIndex);
		if (row != null) {
			return row.getCell(colIndex);
		}
		return null;
	}
	
	public static Cell getOrCreateCell(Sheet sheet, int rowIndex, int colIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		Cell cell = row.getCell(colIndex);
		if (cell == null) {
			cell = row.createCell(colIndex);
		}
		return cell;
	}
	
	public static String formatHyperlink(Book book, int type, String address, String label) {
		if (label == null) {
			label = address;
		}
		final StringBuffer sb  = new StringBuffer(address.length() + label.length() + 128);
		sb.append("<a z.t=\"").append(type).append("\" href=\"")
			.append(address)
			.append("\">")
			.append(label)
			.append("</a>");
		return sb.toString();
	}
	
	public static String formatRichText(Book book, RichTextString rstr, List<int[]> indexes) {
		final String raw = rstr.getString();
		final int runs = rstr.numFormattingRuns();
		if (runs > 0) {
			final int len = rstr.length();
			StringBuffer sb = new StringBuffer(len + runs*96);
			int e = 0;
			boolean prespan = false;
			for(int b = 0, j = 0; j < runs; ++j) {
				final int sbb = sb.length();
				e = rstr.getIndexOfFormattingRun(j);
				if (e > 0) { //first run does NOT start from index 0
					sb.append(raw.substring(b, e));
					if (indexes != null) {
						final int sbe = sb.length();
						indexes.add(new int[] {sbb, sbe});
					}
					if (prespan) {
						sb.append("</span>");
						prespan = false;
					}
				}
				b = e;
				final Font font = getFont(book, rstr, j);
				if (font != null) {
					sb.append("<span style=\"")
						.append(getFontCSSStyle(book, font))
						.append("\">");
					prespan = true;
				}
			}
			if (e < len) {
				final int sbb = sb.length();
				sb.append(raw.substring(e));
				if (indexes != null) {
					final int sbe = sb.length();
					indexes.add(new int[] {sbb, sbe});
				}
				if (prespan) {
					sb.append("</span>");
					prespan = false;
				}
			}
			return sb.toString();
		} else {
			if (indexes != null)
				indexes.add(new int[] {0, raw.length()});
			return raw;
		}
	}
	
	public static String getTextCSSStyle(Book book, Cell cell) {
		final CellStyle style = cell.getCellStyle();
		if (style == null)
			return "";

		final StringBuffer sb = new StringBuffer();
		int textHAlign = getRealAlignment(cell);
		
		switch(textHAlign) {
		case CellStyle.ALIGN_RIGHT:
			sb.append("text-align:").append("right").append(";");
			break;
		case CellStyle.ALIGN_CENTER:
		case CellStyle.ALIGN_CENTER_SELECTION:
			sb.append("text-align:").append("center").append(";");
			break;
		}

		boolean textWrap = style.getWrapText();
		if (textWrap) {
			sb.append("white-space:").append("normal").append(";");
		}/*else{ sb.append("white-space:").append("nowrap").append(";"); }*/

		return sb.toString();
	}

	/** given alignment and cell type, return real alignment */
	//Halignment determined by style alignment, text format and value type  
	public static int getRealAlignment(Cell cell) {
		CellStyle style = cell.getCellStyle();
		int type = cell.getCellType();
		int align = style.getAlignment();
		if (align == CellStyle.ALIGN_GENERAL) {
			final String format = style.getDataFormatString();
			if (format != null && format.startsWith("@")) //a text format
				type = Cell.CELL_TYPE_STRING;
			else if (type == Cell.CELL_TYPE_FORMULA)
				type = cell.getCachedFormulaResultType();
			switch(type) {
			case Cell.CELL_TYPE_BLANK:
				return align;
			case Cell.CELL_TYPE_BOOLEAN:
				return CellStyle.ALIGN_CENTER;
			case Cell.CELL_TYPE_ERROR:
				return CellStyle.ALIGN_CENTER;
			case Cell.CELL_TYPE_NUMERIC:
				return CellStyle.ALIGN_RIGHT;
			case Cell.CELL_TYPE_STRING:
				return CellStyle.ALIGN_LEFT;
			}
		}
		return align;
	}
	
	public static String getFontCSSStyle(Book book, Font font) {
		final StringBuffer sb = new StringBuffer();
		
		String fontName = font.getFontName();
		if (fontName != null) {
			sb.append("font-family:").append(fontName).append(";");
		}
		
		String textColor = BookHelper.getFontHTMLColor(book, font);
		if (BookHelper.AUTO_COLOR.equals(textColor)) {
			textColor = "#000000";
		}
		if (textColor != null) {
			sb.append("color:").append(textColor).append(";");
		}

		final int fontUnderline = font.getUnderline(); 
		final boolean strikeThrough = font.getStrikeout();
		boolean isUnderline = fontUnderline == Font.U_SINGLE || fontUnderline == Font.U_SINGLE_ACCOUNTING;
		if (strikeThrough || isUnderline) {
			sb.append("text-decoration:");
			if (strikeThrough)
				sb.append(" line-through");
			if (isUnderline)	
				sb.append(" underline");
			sb.append(";");
		}

		final short weight = font.getBoldweight();
		final boolean italic = font.getItalic();
		sb.append("font-weight:").append(weight).append(";");
		if (italic)
			sb.append("font-style:").append("italic;");

		final int fontSize = font.getFontHeightInPoints();
		sb.append("font-size:").append(fontSize).append("pt;");
		return sb.toString();
	}
	
	private static Font getFont(Book book, RichTextString rstr, int run) {
		return rstr instanceof HSSFRichTextString ?
				book.getFontAt(((HSSFRichTextString)rstr).getFontOfFormattingRun(run)) :
				((XSSFRichTextString)rstr).getFontOfFormattingRun(run);
		
	}
	
	public static String getFontHTMLColor(Book book, Font font) {
		if (font instanceof XSSFFont) {
			final XSSFFont f = (XSSFFont) font;
			final XSSFColor color = f.getXSSFColor();
			return BookHelper.colorToHTML(book, color);
		} else {
			return indexToHSSFRGB((HSSFWorkbook)book, font.getColor());
		}
	}
	
	/**
	 * Returns the associated #rrggbb HTML color per the given POI Color.
	 * @return the associated #rrggbb HTML color per the given POI Color.
	 */
	public static String colorToHTML(Workbook book, Color color) {
		if (book instanceof HSSFWorkbook) {
			return HSSFColorToHTML((HSSFWorkbook) book, (HSSFColor) color);
		} else {
			return XSSFColorToHTML((XSSFWorkbook) book, (XSSFColor) color);
		}
	}
	
	private static String XSSFColorToHTML(XSSFWorkbook book, XSSFColor color) {
		if (color != null) {
			final CTColor ctcolor = color.getCTColor();
			if (ctcolor.isSetIndexed()) {
				byte[] rgb = IndexedRGB.getRGB(color.getIndexed());
				if (rgb != null) {
					return "#"+ toHex(rgb[0])+ toHex(rgb[1])+ toHex(rgb[2]);
				}
			}
			if (ctcolor.isSetRgb()) {
				byte[] argb = ctcolor.isSetTint() ?
					color.getRgbWithTint() : color.getRgb();
				return argb.length > 3 ? 
					"#"+ toHex(argb[1])+ toHex(argb[2])+ toHex(argb[3])://ignore alpha
					"#"+ toHex(argb[0])+ toHex(argb[1])+ toHex(argb[2]); 
			}
			if (ctcolor.isSetTheme()) {
			    ThemesTable theme = book.getTheme();
			    if (theme != null) {
			    	XSSFColor themecolor = theme.getThemeColor(color.getTheme());
			    	if (themecolor != null) {
			    		return XSSFColorToHTML(book, themecolor);
			    	}
			    }
			}
		}
	    return AUTO_COLOR;
 	}
	
	private static String HSSFColorToHTML(HSSFWorkbook book, HSSFColor color) {
		return color == null ? AUTO_COLOR : indexToHSSFRGB(book, color.getIndex());
	}
	
	private static String indexToHSSFRGB(HSSFWorkbook book, int index) {
		HSSFPalette palette = book.getCustomPalette();
		HSSFColor color = null;
		if (palette != null) {
			color = palette.getColor(index);
		}
		short[] triplet = null;
		if (color != null)
			triplet =  color.getTriplet();
		else {
			final Hashtable<Integer, HSSFColor> colors = HSSFColor.getIndexHash();
			color = colors.get(new Integer(index));
			if (color != null)
				triplet = color.getTriplet();
		}
		return triplet == null ? null : color == HSSFColor.AUTOMATIC.getInstance() ? AUTO_COLOR : 
			"#"+ toHex(triplet[0])+ toHex(triplet[1])+ toHex(triplet[2]);
	}
	
	public static short rgbToIndex(Book book, String color) {
		HSSFPalette palette = ((HSSFWorkbook)book).getCustomPalette();
		short red = Short.parseShort(color.substring(1,3), 16); //red
		short green = Short.parseShort(color.substring(3,5), 16); //green
		short blue = Short.parseShort(color.substring(5), 16); //blue
		byte r = (byte)Math.abs((byte)red);
		byte g = (byte)Math.abs((byte)green);
		byte b = (byte)Math.abs((byte)blue);
		
		HSSFColor pcolor = palette.findColor(r, g, b);
		if (pcolor != null) { //find default palette
			return pcolor.getIndex();
		} else {
			final Hashtable<short[], HSSFColor> colors = HSSFColor.getTripletHash();
			HSSFColor tcolor = colors.get(new short[] {red, green, blue});
			if (tcolor != null)
				return tcolor.getIndex();
			else {
				try {
					HSSFColor ncolor = palette.addColor(r, g, b);
					return ncolor.getIndex();
				} catch (RuntimeException ex) {
					//return similar color if can't add new color to palette
					/**
					 * TODO: find a better solution for fix this issue
					 * 
					 * While there is no space for adding a color into palette
					 * return similar color cause return inexact color
					 */
					return palette.findSimilarColor(red, green, blue).getIndex();
				}
				
			}
		}
	}
	
	public static String toHex(int num) {
		num = num & 0xff;
		final String hex = Integer.toHexString(num);
		return hex.length() == 1 ? "0"+hex : hex; 
	}

	public static Hyperlink getHyperlink(Cell cell) {
		final Hyperlink hyperlink = cell.getHyperlink();
		if (hyperlink != null) {
			return hyperlink; 
		}
		if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			BookHelper.evaluate((Book)cell.getSheet().getWorkbook(), cell);
			final int type = cell.getCachedFormulaResultType();
			if (type != Cell.CELL_TYPE_ERROR) {
				if (cell instanceof HSSFCell)
					return ((HSSFCell)cell).getEvalHyperlink();
				else
					return ((XSSFCell)cell).getEvalHyperlink();
			}
		}
		return null;
	}
	
	public static RichTextString getText(Cell cell) {
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_FORMULA) {
			final Book book = (Book)cell.getSheet().getWorkbook();
			final CellValue cv = BookHelper.evaluate(book, cell);
			cellType = cv.getCellType();
		} else if (cellType == Cell.CELL_TYPE_STRING) {
			return cell.getRichStringCellValue();
		}
	    final String result = new DataFormatter().formatCellValue(cell, cellType);
		return cell instanceof HSSFCell ? new HSSFRichTextString(result) : new XSSFRichTextString(result);
	}
	public static FormatTextImpl getFormatText(Cell cell) {
		int cellType = cell.getCellType();
        final String formatStr = cell.getCellStyle().getDataFormatString();
		if (cellType == Cell.CELL_TYPE_FORMULA) {
			final Book book = (Book)cell.getSheet().getWorkbook();
			final CellValue cv = BookHelper.evaluate(book, cell);
			cellType = cv.getCellType();
//			final Hyperlink hyperlink = cv.getEvalHyperlink();
//			if (hyperlink != null) { //might be HYPERLINK function
//				//do some special format?
//			}
		} 
		if (cellType == Cell.CELL_TYPE_STRING) {
			if ("General".equalsIgnoreCase(formatStr) || "@".equals(formatStr)) {
				return new FormatTextImpl(cell.getRichStringCellValue());
			}
		}
		final CellFormat format = CellFormat.getInstance(formatStr == null ? "" : formatStr);
		return new FormatTextImpl(format.apply(cell));
	}

	//Return simple String form of the cell 
	public static String getCellText(Cell cell) {
		final FormatTextImpl ft = getFormatText(cell);
		if (ft != null) {
			if (ft.isRichTextString()) {
				final RichTextString rstr = ft.getRichTextString();
				return rstr == null ? "" : rstr.getString();
			} else if (ft.isCellFormatResult()) {
				return ft.getCellFormatResult().text;			}
		}
		return "";
	}
	
	public static Object getValueByCellValue(CellValue cellValue) {
		final int cellType = cellValue.getCellType();
		switch(cellType) {
		case Cell.CELL_TYPE_BLANK:
			return "";
		case Cell.CELL_TYPE_BOOLEAN:
			return Boolean.valueOf(cellValue.getBooleanValue());
		case Cell.CELL_TYPE_ERROR:
			return new Byte(cellValue.getErrorValue());
		case Cell.CELL_TYPE_NUMERIC:
			return new Double(cellValue.getNumberValue());
		case Cell.CELL_TYPE_STRING:
			return cellValue.getStringValue();
		}
		throw new UiException("Unknown cell type:"+cellType);
	}
	
	public static String getEditText(Cell cell) {
		final int cellType = cell.getCellType();
		switch(cellType) {
		case Cell.CELL_TYPE_BLANK:
			return "";
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue() ? "TRUE" : "FALSE";
		case Cell.CELL_TYPE_ERROR:
			return ErrorConstants.getText(cell.getErrorCellValue());
		case Cell.CELL_TYPE_FORMULA:
			return "="+cell.getCellFormula();
		case Cell.CELL_TYPE_NUMERIC:
			return NumberToTextConverter.toText(cell.getNumericCellValue());
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue(); 
		default:
			throw new UiException("Unknown cell type:"+cellType);
		}
	}
	
	public static RichTextString getRichEditText(Cell cell) {
		final int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_STRING) {
			return cell.getRichStringCellValue();
		} else {
			return cell.getSheet().getWorkbook().getCreationHelper().createRichTextString(getEditText(cell));
		}
	}

	public static Object getCellValue(Cell cell) {
		final int cellType = cell.getCellType();
		switch(cellType) {
		case Cell.CELL_TYPE_BLANK:
			return "";
		case Cell.CELL_TYPE_BOOLEAN:
			return Boolean.valueOf(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_ERROR:
			return new Byte(cell.getErrorCellValue());
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		case Cell.CELL_TYPE_NUMERIC:
			return new Double(cell.getNumericCellValue());
		case Cell.CELL_TYPE_STRING:
			final RichTextString rtstr = cell.getRichStringCellValue();
			return rtstr == null ? cell.getSheet().getWorkbook().getCreationHelper().createRichTextString("") : rtstr; 
		default:
			throw new UiException("Unknown cell type:"+cellType);
		}
	}
	
	public static Set<Ref>[] removeCell(Sheet sheet, int rowIndex, int colIndex) {
		final Cell cell = getCell(sheet, rowIndex, colIndex);
		return removeCell(cell, true);
	}
	
	public static Set<Ref>[] clearCell(Sheet sheet, int rowIndex, int colIndex) {
		final Cell cell = getCell(sheet, rowIndex, colIndex);
		return clearCell(cell);
	}
	
	private static Set<Ref>[] clearCell(Cell cell) {
		if (cell != null) {
			//remove formula cell and create a blank one
			removeFormula(cell, true);
			//return the affected dependents [0]: last, [1]: all
			final Set<Ref>[] refs = getBothDependents(cell);
			//clear the cell
			//TODO has to clear hyperlink, too
			if (cell != null) {
				cell.setCellValue((String)null);
			}
			return refs;
		}
		return null;
	}
	
	private static Set<Ref>[] removeCell(Cell cell, boolean clearPtgs) {
		if (cell != null) {
			//remove formula cell and create a blank one
			removeFormula(cell, clearPtgs);
			//return the affected dependents [0]: last, [1]: all
			final Set<Ref>[] refs = getBothDependents(cell);
			//remove the cell
			cell.getRow().removeCell(cell);
			return refs;
		}
		return null;
	}
	
	public static Set<Ref>[] setCellValue(Cell cell, Number value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_NUMERIC, value))
			return null;
		//remove formula cell and create a blank one
		removeFormula(cell, true);
		//set value into cell model
		cell.setCellValue(value.doubleValue());
		//return the affected dependents [0]: last, [1]: all
		return getBothDependents(cell);
	}
	
	public static Set<Ref>[] setCellValue(Cell cell, Boolean value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_BOOLEAN, value))
			return null;
		//remove formula cell and create a blank one
		removeFormula(cell, true);
		//set value into cell model
		cell.setCellValue(value.booleanValue());
		//return the affected dependents [0]: last, [1]: all
		return getBothDependents(cell);
	}
	
	public static Set<Ref>[] setCellErrorValue(Cell cell, Byte value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_ERROR, value))
			return null;
		//remove formula cell and crete a blank one
		removeFormula(cell, true);
		//set value into cell model
		cell.setCellErrorValue(value.byteValue());
		//return the affected dependents [0]: last, [1]: all
		return getBothDependents(cell);
	}
	
	public static Set<Ref>[] setCellValue(Cell cell, Date value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_NUMERIC, value))
			return null;
		//remove formula cell and crete a blank one
		removeFormula(cell, true);
		//set value into cell model
		cell.setCellValue(value);
		//notify to update cache
		return getBothDependents(cell); 
	}

	public static Set<Ref>[] setCellFormula(Cell cell, String value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_FORMULA, value))
			return null;
		//remove formula cell and crete a blank one
		removeFormula(cell, true);
		//set value into cell model
		cell.setCellFormula(value);
		//notify to update cache
		return getBothDependents(cell); 
	}
	
	public static Set<Ref>[] setCellHyperlink(Cell cell, int linkTarget, String address) {
		Hyperlink hlink = cell.getHyperlink();
		if (hlink == null) {
			Workbook wb = cell.getSheet().getWorkbook();
			CreationHelper createHelper = wb.getCreationHelper();
			Hyperlink link = createHelper.createHyperlink(linkTarget);
			link.setAddress(address);
			
			cell.setHyperlink(link);
		} else {
			
			if (sameHyperlink(cell, hlink, linkTarget, address))
				return null;

			hlink.setAddress(address);
		}
		//notify to update cache
		return getBothDependents(cell); 
	}
	
	private static boolean sameHyperlink(Cell cell, Hyperlink hlink,
			int linkType, String address) {		
		return hlink.getType() == linkType && hlink.getAddress() == address;
	}

	public static Set<Ref>[] setCellValue(Cell cell, String value) {
		return setCellValue(cell, value == null ? null : cell.getSheet().getWorkbook().getCreationHelper().createRichTextString(value));
	}
	
	public static Set<Ref>[] setCellValue(Cell cell, RichTextString value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_STRING, value))
			return null;
		//remove formula cell and crete a blank one
		removeFormula(cell, true);
		//set value into cell model
		cell.setCellValue(value);
		//notify to update cache
		return getBothDependents(cell); 
	}
	
	private static boolean sameTypeAndValue(Cell cell, int type, Object value) {
		if (cell.getCellType() != type) return false;
		final Object cellValue = getCellValue(cell);
		return cellValue == null ? cellValue == value : cellValue.equals(value);
	}
	
	//[0]:last, [1]:all
	private static Set<Ref>[] getBothDependents(Cell cell) {
		final Sheet sheet = cell.getSheet();
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		
		//clear/update formula cache
		book.getFormulaEvaluator().notifySetFormula(cell); 
		//get affected dependents(last, all)
		final int row = cell.getRowIndex();
		final int col = cell.getColumnIndex();
		Set<Ref>[] refs = ((RefSheetImpl)refSheet).getBothDependents(row, col);
		//no dependent but myself is a formula cell
		if (refs[0].isEmpty() && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			final Ref srcRef = refSheet.getRef(row, col, row, col);
			if (srcRef != null) {
				refs[0].add(srcRef);
				refs[1].add(srcRef);
			}
		}
		return refs;
	}
	
	//formula cell -> non-formula cell
	private static void removeFormula(Cell cell, boolean clearFormula) {
		//Was formula, shall maintain the formula cache and dependency 
		if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			//clear the formula reference dependency
			final int rowIndex = cell.getRowIndex(); 
			final int colIndex = cell.getColumnIndex();
			final Sheet sheet = cell.getSheet();
			final Book book = (Book) sheet.getWorkbook();
			final RefSheet refSheet = getRefSheet(book, sheet);
			final Ref ref = refSheet.getRef(rowIndex, colIndex, rowIndex, colIndex);
			//bug 60: Sort function when seelct multiple columns cause NullPointer Exception
			//ref could be null if never evaluated or cleared if an orphan reference during operation
			if (ref != null) { 
				ref.removeAllPrecedents();
			}
			
			//remove formula from the cell
			if (clearFormula) {
				cell.setCellFormula(null);
			}
			
			//clear formula cache
			book.getFormulaEvaluator().notifySetFormula(cell); 
		}
	}
	
	public static Object[] editTextToValue(String txt) {
		if (txt != null) {
			if (txt.startsWith("=")) {
				if (txt.trim().length() > 1) {
					return new Object[] {new Integer(Cell.CELL_TYPE_FORMULA), txt.substring(1)}; //formula 
				} else {
					return new Object[] {new Integer(Cell.CELL_TYPE_STRING), txt}; //string
				}
			} else if ("true".equalsIgnoreCase(txt) || "false".equalsIgnoreCase(txt)) {
				return new Object[] {new Integer(Cell.CELL_TYPE_BOOLEAN), Boolean.valueOf(txt)}; //boolean
			} else if (txt.startsWith("#")) { //might be an error
				final byte err = getErrorCode(txt);
				return err < 0 ? 
					new Object[] {new Integer(Cell.CELL_TYPE_STRING), txt}: //string
					new Object[] {new Integer(Cell.CELL_TYPE_ERROR), new Byte(err)}; //error
			} else {
				try {
					final Double val = Double.parseDouble(txt);
					return new Object[] {new Integer(Cell.CELL_TYPE_NUMERIC), val}; //double
				} catch (NumberFormatException ex) {
					final Object[] results = parseToDate(txt); 
					if (results[0] instanceof String) { 
						return new Object[] {new Integer(Cell.CELL_TYPE_STRING), results[0]}; //string
					} else { //if (result[0] instanceof Date)
						return new Object[] {new Integer(Cell.CELL_TYPE_NUMERIC), results[0], results[1]}; //date with format
					}
				}
			}
		}
		return null;
	}

	private static byte getErrorCode(String errString) {
		if ("#NULL!".equals(errString))
			return ErrorConstants.ERROR_NULL;
		if ("#DIV/0!".equals(errString))
			return ErrorConstants.ERROR_DIV_0;
		if ("#VALUE!".equals(errString))
			return ErrorConstants.ERROR_VALUE;
		if ("#REF!".equals(errString))
			return ErrorConstants.ERROR_REF;
		if ("#NAME?".equals(errString))
			return ErrorConstants.ERROR_NAME;
		if ("#NUM!".equals(errString))
			return ErrorConstants.ERROR_NUM; 
		if ("#N/A".equals(errString))
			return ErrorConstants.ERROR_NA;
		return -1;
	}

	//0: value 1: format
	public static Object[] parseToDate(String txt) {
		//TODO when BookHelper refactor to singleton, DateInputMask can be a private variable 
		return new DateInputMask().parseDateInput(txt);
	}
	
	public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		return row;
	}
	
	private static CellStyle prepareCellStyle(CellStyle srcStyle, Cell dstCell, int pasteType) {
		//TODO now assume same workbook in copying CellStyle.
		if ((pasteType & Range.PASTE_FORMATS) == BookHelper.INNERPASTE_NUMBER_FORMATS) { //number format only
			final CellStyle style = dstCell.getSheet().getWorkbook().createCellStyle();
			final CellStyle dstStyle = dstCell.getCellStyle();
			style.cloneStyleFrom(dstStyle);
			final short fmt = srcStyle.getDataFormat();
			style.setDataFormat(fmt);
			return style;
		}
		if ((pasteType & BookHelper.INNERPASTE_BORDERS) == 0) { //no border
			final CellStyle style = dstCell.getSheet().getWorkbook().createCellStyle();
			final short borderLeft = style.getBorderLeft();
			final short borderTop = style.getBorderTop();
			final short borderRight = style.getBorderRight();
			final short borderBottom = style.getBorderBottom();
			final short borderLeftColor = style.getLeftBorderColor();
			final short borderTopColor = style.getTopBorderColor();
			final short borderRightColor = style.getRightBorderColor();
			final short borderBottomColor = style.getBottomBorderColor();
			style.cloneStyleFrom(srcStyle);
			style.setBorderLeft(borderLeft );
			style.setBorderTop(borderTop);
			style.setBorderRight(borderRight);
			style.setBorderBottom(borderBottom);
			style.setLeftBorderColor(borderLeftColor);
			style.setTopBorderColor(borderTopColor);
			style.setRightBorderColor(borderRightColor);
			style.setBottomBorderColor(borderBottomColor);
			return style;
		}
		return srcStyle;
	}
	
	private static void copyComment(Cell srcCell, Cell dstCell) {
		final Comment srcComment = srcCell.getCellComment();
		Comment dstComment = dstCell.getCellComment();
		if (srcComment != null) {
			if (dstComment == null) {
				final Sheet dstSheet = dstCell.getSheet();
				final Workbook dstBook = dstSheet.getWorkbook();
				final CreationHelper dstFactory = dstBook.getCreationHelper();
				final Drawing drawing = dstSheet.createDrawingPatriarch();
				final ClientAnchor anchor = dstFactory.createClientAnchor();
				dstComment = drawing.createCellComment(anchor);
			}
			dstComment.setString(srcComment.getString());
			dstComment.setAuthor(srcComment.getAuthor());
			dstComment.setVisible(srcComment.isVisible());
			dstCell.setCellComment(dstComment);
		} else { //srcComment is null
			if (dstComment != null) {
				dstCell.removeCellComment();
			}
		}
	}
	
	private static void copyValidation(Cell srcCell, Cell dstCell) {
		//TODO now assume same workbook in copying CellStyle.
		final Sheet srcSheet = srcCell.getSheet();
		final Sheet dstSheet = dstCell.getSheet();
		
		final int srcRow = srcCell.getRowIndex();
		final int srcCol = srcCell.getColumnIndex();
		final int dstRow = dstCell.getRowIndex();
		final int dstCol = dstCell.getColumnIndex();
		
		final List<? extends DataValidation> dataValidations= BookHelper.getDataValidations(srcSheet);
		if (dstSheet.equals(srcSheet)) {
			for(DataValidation dataValidation : dataValidations) {
				CellRangeAddressList addrList = dataValidation.getRegions();
				boolean srcInRange = false;
				boolean dstInRange = false;
				for(int j = addrList.countRanges(); --j >= 0;) {
					final CellRangeAddress addr = addrList.getCellRangeAddress(j);
					if (!srcInRange) {
						srcInRange = addr.isInRange(srcRow, srcCol);
					}
					if (!dstInRange) {
						dstInRange = addr.isInRange(dstRow, dstCol);
					}
					if (srcInRange && dstInRange) { //no need to copy
						break;
					}
				}
				if (!srcInRange) { //this validation is not associated to source cell 
					continue;
				}
				if (!dstInRange) { //so we shall copy this data validation to dst cell
					dataValidation.getRegions().addCellRangeAddress(dstRow, dstCol, dstRow, dstCol);
				}
			}
		} else {
			final DataValidationHelper helper = dstSheet.getDataValidationHelper();
			for(DataValidation dataValidation : dataValidations) {
				CellRangeAddressList addrList = dataValidation.getRegions();
				boolean srcInRange = false;
				for(int j = addrList.countRanges(); --j >= 0;) {
					final CellRangeAddress addr = addrList.getCellRangeAddress(j);
					if (!srcInRange) {
						srcInRange = addr.isInRange(srcRow, srcCol);
					}
					if (srcInRange) {
						break;
					}
				}
				if (!srcInRange) { //this validation is not associated to source cell 
					continue;
				}
				//so we shall copy this data validation to dst cell
				final DataValidationConstraint constraint = BookHelper.getConstraint(dataValidation);
				DataValidation dstDataValidation = BookHelper.getDataValidationByConstraint(constraint, getDataValidations(dstSheet));
				if (dstDataValidation == null) {
					final CellRangeAddressList dstAddrList = new CellRangeAddressList(dstRow, dstCol, dstRow, dstCol);
					dstDataValidation = helper.createValidation(constraint, dstAddrList);
					dstSheet.addValidationData(dstDataValidation);
				} else {
					CellRangeAddressList dstAddrList = dstDataValidation.getRegions();
					dstAddrList.addCellRangeAddress(dstRow, dstCol, dstRow, dstCol);
				}
			}
		}
	}
	private static List<? extends DataValidation> getDataValidations(Sheet sheet) {
		if (sheet instanceof HSSFSheet) {
			return ((HSSFSheetImpl)sheet).getDataValidations();
		} else {
			return ((XSSFSheet)sheet).getDataValidations();
		}
	}
	private static DataValidation getDataValidationByConstraint(DataValidationConstraint constraint, List<? extends DataValidation> dataValidations) {
		for (DataValidation dataValidation : dataValidations) {
			if (constraint.equals(getConstraint(dataValidation)))
				return dataValidation;
		}
		return null;
	}
	private static DataValidationConstraint getConstraint(DataValidation dataValidation) {
		if (dataValidation instanceof HSSFDataValidation) {
			return ((HSSFDataValidation) dataValidation).getConstraint();
		} else {
			return ((XSSFDataValidation) dataValidation).getValidationConstraint();
		}
	}
	
	//[0]:last, [1]:all
	public static Set<Ref>[] copyCell(Cell srcCell, Sheet sheet, int rowIndex, int colIndex, int pasteType, int pasteOp) {
		//TODO not handle pastType == pasteValidation and pasteOp(assume none)
		final Cell dstCell = getOrCreateCell(sheet, rowIndex, colIndex);
		return copyCell(srcCell, dstCell, pasteType, pasteOp);
	}
	
	public static Set<Ref>[] copyCell(Cell srcCell, Cell dstCell, int pasteType, int pasteOp) {
		//paste cell formats
		if ((pasteType & BookHelper.INNERPASTE_FORMATS) != 0) {
			dstCell.setCellStyle(prepareCellStyle(srcCell.getCellStyle(), dstCell, pasteType));
		}
		
		//paste comment
		if ((pasteType & BookHelper.INNERPASTE_COMMENTS) != 0) {
			copyComment(srcCell, dstCell);
		}
		
		//paste validation
		if ((pasteType & BookHelper.INNERPASTE_VALIDATION) != 0) {
			copyValidation(srcCell, dstCell);
		}
		
		//paste value and formula
		if ((pasteType & BookHelper.INNERPASTE_VALUES_AND_FORMULAS) != 0) {
			final int cellType = srcCell.getCellType(); 
			switch(cellType) {
			case Cell.CELL_TYPE_BOOLEAN:
				return setCellValue(dstCell, srcCell.getBooleanCellValue());
			case Cell.CELL_TYPE_ERROR:
				return setCellErrorValue(dstCell, srcCell.getErrorCellValue());
	        case Cell.CELL_TYPE_NUMERIC:
	        	return setCellValue(dstCell, srcCell.getNumericCellValue());
	        case Cell.CELL_TYPE_STRING:
	        	return setCellValue(dstCell, srcCell.getRichStringCellValue());
	        case Cell.CELL_TYPE_BLANK:
	        	return setCellValue(dstCell, (RichTextString) null);
	        case Cell.CELL_TYPE_FORMULA:
	        	if ((pasteType & BookHelper.INNERPASTE_FORMULAS) != 0) { //copy formula
	        		return copyCellFormula(dstCell, srcCell);
	        	} else { //copy evaluated value only
	        		final Book book = (Book) srcCell.getSheet().getWorkbook();
	        		final CellValue cv = evaluate(book, srcCell);
	        		return setCellValueByCellValue(dstCell, cv, pasteType, pasteOp);
	        	}
			default:
				throw new UiException("Unknown cell type:"+cellType);
			}
		}
		
		return null;
	}
	
	
	private static Set<Ref>[] setCellValueByCellValue(Cell dstCell, CellValue cv, int pasteType, int pasteOp) {
		final int cellType = cv.getCellType();
		switch(cellType) {
		case Cell.CELL_TYPE_BOOLEAN:
			return setCellValue(dstCell, cv.getBooleanValue());
		case Cell.CELL_TYPE_ERROR:
			return setCellErrorValue(dstCell, cv.getErrorValue());
        case Cell.CELL_TYPE_NUMERIC:
        	return setCellValue(dstCell, cv.getNumberValue());
        case Cell.CELL_TYPE_STRING:
        	return setCellValue(dstCell, cv.getStringValue());
        case Cell.CELL_TYPE_BLANK:
        	return setCellValue(dstCell, (RichTextString) null);
		default:
			throw new UiException("Unknown cell type in CellValue:"+cellType);
		}
	}
	
	//[0]:last, [1]:all
	private static Set<Ref>[] copyCellFormula(Cell dstCell, Cell srcCell) {
		//remove formula cell and create a blank one
		removeFormula(dstCell, true);
		
		//set value into cell model
		final Ptg[] dstPtgs = offsetPtgs(dstCell, srcCell);
		setCellPtgs(dstCell, dstPtgs);
		evaluate((Book)dstCell.getSheet().getWorkbook(), dstCell);
		
		//notify to update cache
		return getBothDependents(dstCell);
	}
	
	private static void setCellPtgs(Cell cell, Ptg[] ptgs) {
        if (cell instanceof HSSFCell) {
        	setHSSFCellPtgs((HSSFCell)cell, ptgs);
        } else {
        	setXSSFCellPtgs((XSSFCell)cell, ptgs);
        }
	}
	private static void setHSSFCellPtgs(HSSFCell cell, Ptg[] ptgs) {
    	//tricky! must be after the dummyCell construction or the under aggregate record will not 
    	//be consistent in sheet and cell
        cell.setCellType(Cell.CELL_TYPE_FORMULA); 
    	FormulaRecordAggregate agg = (FormulaRecordAggregate) new HSSFCellHelper((HSSFCell)cell).getCellValueRecord();
    	FormulaRecord frec = agg.getFormulaRecord();
    	frec.setOptions((short) 2);
    	frec.setValue(0);

    	//only set to default if there is no extended format index already set
    	if (agg.getXFIndex() == (short)0) {
    		agg.setXFIndex((short) 0x0f);
    	}
    	agg.setParsedExpression(ptgs);
	}
	private static void setXSSFCellPtgs(XSSFCell cell, Ptg[] ptgs) {
		XSSFWorkbook book = (XSSFWorkbook) cell.getSheet().getWorkbook();
        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.create(book);
        final String formula = FormulaRenderer.toFormulaString(fpb, ptgs);
        if (formula != null && formula.length() > 0) {
        	cell.setCellFormula(formula);
        }
	}
	public static Ptg[] getCellPtgs(Cell cell) {
		if (cell instanceof HSSFCell) {
			return getHSSFPtgs((HSSFCell)cell);
		} else if (cell instanceof XSSFCell) {
			final String formula = cell.getCellFormula();
			final XSSFSheet sheet = (XSSFSheet) cell.getSheet();
			final XSSFWorkbook wb = (XSSFWorkbook) sheet.getWorkbook();
			final FormulaParsingWorkbook fpb = XSSFEvaluationWorkbook.create(wb); 
			return FormulaParser.parse(formula, fpb, FormulaType.CELL, wb.getSheetIndex(sheet));
		}
		throw new UiException("Unknown cell class:"+cell);
	}
	
	private static Ptg[] getHSSFPtgs(HSSFCell cell) {
		CellValueRecordInterface vr = new HSSFCellHelper(cell).getCellValueRecord();
		if (!(vr instanceof FormulaRecordAggregate)) {
			throw new IllegalArgumentException("Not a formula cell");
		}
		FormulaRecordAggregate fra = (FormulaRecordAggregate) vr;
		return fra.getFormulaTokens();
	}
	
	private static Ptg[] offsetPtgs(Cell dstCell, Cell srcCell) {
		final Sheet srcSheet = srcCell.getSheet();
		final int srcRow = srcCell.getRowIndex();
		final int srcCol = srcCell.getColumnIndex();
		
		final Sheet dstSheet = dstCell.getSheet();
		final int dstRow = dstCell.getRowIndex();
		final int dstCol = dstCell.getColumnIndex();
		
		final int offRow = dstRow - srcRow;
		final int offCol = dstCol - srcCol;

		return offsetPtgs(srcCell, srcSheet, dstSheet, -1, offRow, -1, offCol);
	}
	
	private static Ptg[] offsetPtgs(Cell srcCell, Sheet srcSheet, Sheet dstSheet, int startRow, int offRow, int startCol, int offCol) {
		final Ptg[] srcPtgs = getCellPtgs(srcCell);
		final int ptglen = srcPtgs.length;
		final Ptg[] dstPtgs = new Ptg[ptglen];
		
		final SpreadsheetVersion ver = ((Book)dstSheet.getWorkbook()).getSpreadsheetVersion();
		for(int j = 0; j < ptglen; ++j) {
			final Ptg srcPtg = srcPtgs[j];
			final Ptg dstPtg = offsetPtg(srcPtg, offRow, offCol, ver);
			dstPtgs[j] = dstPtg;
		}
		return dstPtgs;
	}
	
	private static Ptg rptgValidate(RefPtgBase rptg, int row, int col, SpreadsheetVersion ver) {
		if (col < 0 || col > ver.getLastColumnIndex()
			|| row < 0 || row > ver.getLastRowIndex()) {
			return createDeletedRef(rptg);
		}
		return null;
	}
	
	private static Ptg aptgValidate(AreaPtgBase aptg , int row1, int row2, int col1, int col2, SpreadsheetVersion ver) {
		if (row1 < 0 || row2 > ver.getLastRowIndex()
			|| col1 < 0 || col2 > ver.getLastColumnIndex()) {
			return createDeletedRef(aptg);
		}
		return null;
	}

	private static Ptg createDeletedRef(Ptg ptg) {
		return PtgShifter.createDeletedRef(ptg);
	}

	private static Ptg offsetPtg(Ptg ptg, int offRow, int offCol, SpreadsheetVersion ver) {
		if(ptg instanceof RefPtgBase) {
			final RefPtgBase rptg = (RefPtgBase)ptg;
			return rptgSetRowCol(rptg, offRow, offCol, ver);
		}
		if(ptg instanceof AreaPtgBase) {
			final AreaPtgBase aptg = (AreaPtgBase) ptg;
			return aptgSetRowCol(aptg, offRow, offCol, ver);
		}
		return ptg;
	}

	private static Ptg rptgSetRowCol(RefPtgBase ptg, int nrow, int ncol, SpreadsheetVersion ver) {
		final int row = ptg.getRow() + nrow;
		final int col = ptg.getColumn() + ncol;
		final Ptg xptg = rptgValidate(ptg, row, col, ver);
		if (xptg == null) {
			ptg.setRow(row);
			ptg.setColumn(col);
			return ptg;
		} else
			return xptg;
	}
	
	private static Ptg aptgSetRowCol(AreaPtgBase ptg, int nrow, int ncol, SpreadsheetVersion ver) {
		final int row1 = ptg.getFirstRow() + nrow;
		final int col1 = ptg.getFirstColumn() + ncol;
		final int row2 = ptg.getLastRow() + nrow;
		final int col2 = ptg.getLastColumn() + ncol;
		final Ptg xptg = aptgValidate(ptg, row1, row2, col1, col2, ver);
		if (xptg == null) {
			ptg.setFirstRow(row1);
			ptg.setFirstColumn(col1);
			ptg.setLastRow(row2);
			ptg.setLastColumn(col2);
			return ptg;
		} else {
			return xptg;
		}
	}
	
	public CellStyle findCellStyle(Book book, 
		short dataFormat, short fontIndex, boolean hidden, boolean locked, 
		short alignment, boolean wrapText, short valign, short rotation,
		short indention, 
		short borderLeft, short borderTop, short borderRight, short borderBottom, 
		short borderLeftColor, short borderTopColor, short borderRightColor, short borderBottomColor,
		short fillPattern, short fillBackColor, short fillForeColor) {
		for(short j = 0, len = book.getNumCellStyles(); j < len; ++j) {
			CellStyle style = book.getCellStyleAt(j);
			if (style.getDataFormat() != dataFormat)
				continue;
			if (style.getFontIndex() != fontIndex)
				continue;
			if (style.getFillForegroundColor() != fillForeColor)
				continue;
			if (style.getHidden() != hidden)
				continue;
			if (style.getLocked() != locked)
				continue;
			if (style.getAlignment() != alignment)
				continue;
			if (style.getWrapText() != wrapText)
				continue;
			if (style.getVerticalAlignment() != valign)
				continue;
			if (style.getIndention() != indention)
				continue;
			if (style.getBorderLeft() != borderLeft)
				continue;
			if (style.getBorderTop() != borderTop)
				continue;
			if (style.getBorderRight() != borderRight)
				continue;
			if (style.getBorderBottom() != borderBottom)
				continue;
			if (style.getLeftBorderColor() != borderLeftColor)
				continue;
			if (style.getTopBorderColor() != borderTopColor)
				continue;
			if (style.getRightBorderColor() != borderRightColor)
				continue;
			if (style.getBottomBorderColor() != borderBottomColor)
				continue;
			if (style.getFillPattern() != fillPattern)
				continue;
			if (style.getFillBackgroundColor() != fillBackColor)
				continue;
			if (style.getRotation() != rotation)
				continue;
			return style; //found it!
		}
		return null;
	}
	
	//TODO interface for SheetCtrl?
	public static ChangeInfo insertRows(Sheet sheet, int startRow, int num, int copyOrigin) {
		if (sheet instanceof HSSFSheet) {
			return insertHSSFRows(sheet, startRow, num, copyOrigin);
		} else {
			return insertXSSFRows(sheet, startRow, num, copyOrigin);
		}
	}
	
	private static ChangeInfo insertXSSFRows(Sheet sheet, int startRow, int num, int copyOrigin) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertRows(startRow, num);
		final int lastRowNum = sheet.getLastRowNum();
		if (startRow > lastRowNum) {
			return null;
		}
		final List<CellRangeAddress[]> shiftedRanges = ((XSSFSheetImpl)sheet).shiftRowsOnly(startRow, lastRowNum, num, true, false, true, false, copyOrigin);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, startRow, maxrow, num, 0, maxcol, 0);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	private static ChangeInfo insertHSSFRows(Sheet sheet, int startRow, int num, int copyOrigin) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertRows(startRow, num);
		final int lastRowNum = sheet.getLastRowNum();
		if (startRow > lastRowNum) {
			return null;
		}
		final List<CellRangeAddress[]> shiftedRanges = ((HSSFSheetImpl)sheet).shiftRowsOnly(startRow, lastRowNum, num, true, false, true, false, copyOrigin);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, startRow, maxrow, num, 0, maxcol, 0);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	//TODO interface for SheetCtrl?
	public static ChangeInfo deleteRows(Sheet sheet, int startRow, int num) {
		if (sheet instanceof HSSFSheet) {
			return deleteHSSFRows(sheet, startRow, num);
		} else {
			return deleteXSSFRows(sheet, startRow, num);
		}
	}
	private static ChangeInfo deleteHSSFRows(Sheet sheet, int startRow, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteRows(startRow, num);
		final int lastRowNum = sheet.getLastRowNum();
		final int startRow0 = startRow + num;
		if (startRow > lastRowNum) {
			return null;
		}
		final List<CellRangeAddress[]> shiftedRanges = ((HSSFSheetImpl)sheet).shiftRowsOnly(startRow0, lastRowNum, -num, true, false, true, true, Range.FORMAT_NONE);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, startRow0, maxrow, -num, 0, maxcol, 0);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	private static ChangeInfo deleteXSSFRows(Sheet sheet, int startRow, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteRows(startRow, num);
		final int lastRowNum = sheet.getLastRowNum();
		final int startRow0 = startRow + num;
		if (startRow > lastRowNum) {
			return null;
		}
		final List<CellRangeAddress[]> shiftedRanges = ((XSSFSheetImpl)sheet).shiftRowsOnly(startRow0, lastRowNum, -num, true, false, true, true, Range.FORMAT_NONE);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, startRow0, maxrow, -num, 0, maxcol, 0);
		
		return new ChangeInfo(last, all, changeMerges);
	}

	//TODO interface for SheetCtrl?
	public static ChangeInfo insertColumns(Sheet sheet, int startCol, int num, int copyOrigin) {
		if (sheet instanceof HSSFSheet) {
			return insertHSSFColumns(sheet, startCol, num, copyOrigin);
		} else {
			return insertXSSFColumns(sheet, startCol, num, copyOrigin);
		}
	}
			
	private static ChangeInfo insertHSSFColumns(Sheet sheet, int startCol, int num, int copyOrigin) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertColumns(startCol, num);
		final List<CellRangeAddress[]> shiftedRanges = ((HSSFSheetImpl)sheet).shiftColumnsOnly(startCol, -1, num, true, false, true, false, copyOrigin);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, 0, maxrow, 0, startCol, maxcol, num);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	private static ChangeInfo insertXSSFColumns(Sheet sheet, int startCol, int num, int copyOrigin) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertColumns(startCol, num);
		final List<CellRangeAddress[]> shiftedRanges = ((XSSFSheetImpl)sheet).shiftColumnsOnly(startCol, -1, num, true, false, true, false, copyOrigin);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, 0, maxrow, 0, startCol, maxcol, num);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	//TODO SheetCtrl interface?
	public static ChangeInfo deleteColumns(Sheet sheet, int startCol, int num) {
		if (sheet instanceof HSSFSheet) {
			return deleteHSSFColumns(sheet, startCol, num);
		} else {
			return deleteXSSFColumns(sheet, startCol, num);
		}
	}
	
	public static ChangeInfo deleteHSSFColumns(Sheet sheet, int startCol, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteColumns(startCol, num);
		final int startCol0 = startCol + num;
		final List<CellRangeAddress[]> shiftedRanges = ((HSSFSheetImpl)sheet).shiftColumnsOnly(startCol0, -1, -num, true, false, true, true, Range.FORMAT_NONE);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, 0, maxrow, 0, startCol0, maxcol, -num);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	public static ChangeInfo deleteXSSFColumns(Sheet sheet, int startCol, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteColumns(startCol, num);
		final int startCol0 = startCol + num;
		final List<CellRangeAddress[]> shiftedRanges = ((XSSFSheetImpl)sheet).shiftColumnsOnly(startCol0, -1, -num, true, false, true, true, Range.FORMAT_NONE);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		shiftFormulas(all, sheet, 0, maxrow, 0, startCol0, maxcol, -num);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	public static ChangeInfo insertRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean horizontal, int copyRightBelow) {
		if (sheet instanceof HSSFSheet) {
			return insertHSSFRange(sheet, tRow, lCol, bRow, rCol, horizontal, copyRightBelow);
		} else {
			return insertXSSFRange(sheet, tRow, lCol, bRow, rCol, horizontal, copyRightBelow);
		}
	}
	
	private static ChangeInfo insertHSSFRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean horizontal, int copyRightBelow) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertRange(tRow, lCol, bRow, rCol, horizontal);
		final int num = horizontal ? rCol - lCol + 1 : bRow - tRow + 1;
		final List<CellRangeAddress[]> shiftedRanges = horizontal ? 
			((HSSFSheetImpl)sheet).shiftColumnsRange(lCol, -1, num, tRow, bRow, true, false, true, false, copyRightBelow):
			((HSSFSheetImpl)sheet).shiftRowsRange(tRow, -1, num, lCol, rCol, true, false, true, false, copyRightBelow);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		if (horizontal) {
			final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
			all.add(new AreaRefImpl(tRow, lCol, bRow, maxcol, refSheet));
			shiftFormulas(all, sheet, tRow, bRow, 0, lCol, maxcol, num);
		} else {
			final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
			all.add(new AreaRefImpl(tRow, lCol, maxrow, rCol, refSheet));
			shiftFormulas(all, sheet, tRow, maxrow, num, lCol, rCol, 0);
		}
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	private static ChangeInfo insertXSSFRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean horizontal, int copyRightBelow) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertRange(tRow, lCol, bRow, rCol, horizontal);
		final int num = horizontal ? rCol - lCol + 1 : bRow - tRow + 1;
		final List<CellRangeAddress[]> shiftedRanges = horizontal ? 
			((XSSFSheetImpl)sheet).shiftColumnsRange(lCol, -1, num, tRow, bRow, true, false, true, false, copyRightBelow):
			((XSSFSheetImpl)sheet).shiftRowsRange(tRow, -1, num, lCol, rCol, true, false, true, false, copyRightBelow);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		if (horizontal) {
			final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
			all.add(new AreaRefImpl(tRow, lCol, bRow, maxcol, refSheet));
			shiftFormulas(all, sheet, tRow, bRow, 0, lCol, maxcol, num);
		} else {
			final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
			all.add(new AreaRefImpl(tRow, lCol, maxrow, rCol, refSheet));
			shiftFormulas(all, sheet, tRow, maxrow, num, lCol, rCol, 0);
		}
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	public static ChangeInfo deleteRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean horizontal) {
		if (sheet instanceof HSSFSheet) {
			return deleteHSSFRange(sheet, tRow, lCol, bRow, rCol, horizontal);
		} else {
			return deleteXSSFRange(sheet, tRow, lCol, bRow, rCol, horizontal);
		}
	}
	
	private static ChangeInfo deleteHSSFRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean horizontal) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteRange(tRow, lCol, bRow, rCol, horizontal);
		final int num = horizontal ? rCol - lCol + 1 : bRow - tRow + 1;
		final List<CellRangeAddress[]> shiftedRanges = horizontal ? 
			((HSSFSheetImpl)sheet).shiftColumnsRange(rCol + 1, -1, -num, tRow, bRow, true, false, true, true, Range.FORMAT_NONE):
			((HSSFSheetImpl)sheet).shiftRowsRange(bRow + 1, -1, -num, lCol, rCol, true, false, true, true, Range.FORMAT_NONE);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		
		if (horizontal) {
			final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
			all.add(new AreaRefImpl(tRow, lCol, bRow, maxcol, refSheet));
			shiftFormulas(all, sheet, tRow, bRow, 0, rCol + 1, maxcol, -num);
		} else {
			final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
			all.add(new AreaRefImpl(tRow, lCol, maxrow, rCol, refSheet));
			shiftFormulas(all, sheet, bRow + 1, maxrow, -num, lCol, rCol, 0);
		}
		
		return new ChangeInfo(last, all, changeMerges);
	}

	private static ChangeInfo deleteXSSFRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean horizontal) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteRange(tRow, lCol, bRow, rCol, horizontal);
		final int num = horizontal ? rCol - lCol + 1 : bRow - tRow + 1;
		final List<CellRangeAddress[]> shiftedRanges = horizontal ? 
			((XSSFSheetImpl)sheet).shiftColumnsRange(rCol + 1, -1, -num, tRow, bRow, true, false, true, true, Range.FORMAT_NONE):
			((XSSFSheetImpl)sheet).shiftRowsRange(bRow + 1, -1, -num, lCol, rCol, true, false, true, true, Range.FORMAT_NONE);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		
		if (horizontal) {
			final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
			all.add(new AreaRefImpl(tRow, lCol, bRow, maxcol, refSheet));
			shiftFormulas(all, sheet, tRow, bRow, 0, rCol + 1, maxcol, -num);
		} else {
			final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
			all.add(new AreaRefImpl(tRow, lCol, maxrow, rCol, refSheet));
			shiftFormulas(all, sheet, bRow + 1, maxrow, -num, lCol, rCol, 0);
		}
		
		return new ChangeInfo(last, all, changeMerges);
	}

	//TODO as SheetCtrl interface
	public static ChangeInfo moveRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, int nRow, int nCol) {
		if (sheet instanceof HSSFSheet) {
			return moveHSSFRange(sheet, tRow, lCol, bRow, rCol, nRow, nCol);
		} else {
			return moveXSSFRange(sheet, tRow, lCol, bRow, rCol, nRow, nCol);
		}
	}
		
	private static ChangeInfo moveHSSFRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, int nRow, int nCol) {
		if (nRow == 0 && nCol == 0) { //nothing to do!
			return null;
		}
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.moveRange(tRow, lCol, bRow, rCol, nRow, nCol);
		final List<CellRangeAddress[]> shiftedRanges = nCol != 0 && nRow == 0 ? 
			((HSSFSheetImpl)sheet).shiftColumnsRange(lCol, rCol, nCol, tRow, bRow, true, false, true, false, Range.FORMAT_NONE):
			nCol == 0 && nRow != 0 ?
			((HSSFSheetImpl)sheet).shiftRowsRange(tRow, bRow, nRow, lCol, rCol, true, false, true, false, Range.FORMAT_NONE):
			((HSSFSheetImpl)sheet).shiftBothRange(tRow, bRow, nRow, lCol, rCol, nCol, true); //nCol != 0 && nRow != 0
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		shiftFormulas(all, sheet, tRow, bRow, nRow, lCol, rCol, nCol);
		all.add(new AreaRefImpl(tRow, lCol, bRow, rCol, refSheet)); //original selection
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		int ntRow = Math.max(0, tRow + nRow);
		int nlCol = Math.max(0, lCol + nCol);
		int nbRow = Math.min(maxrow, bRow + nRow);
		int nrCol = Math.min(maxcol, rCol + nCol);
		if (ntRow <= nbRow && nlCol <= nrCol)
		all.add(new AreaRefImpl(ntRow, nlCol, nbRow, nrCol, refSheet));
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	private static ChangeInfo moveXSSFRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol, int nRow, int nCol) {
		if (nRow == 0 && nCol == 0) { //nothing to do!
			return null;
		}
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.moveRange(tRow, lCol, bRow, rCol, nRow, nCol);
		final List<CellRangeAddress[]> shiftedRanges = nCol != 0 && nRow == 0 ? 
			((XSSFSheetImpl)sheet).shiftColumnsRange(lCol, rCol, nCol, tRow, bRow, true, false, true, false, Range.FORMAT_NONE):
			nCol == 0 && nRow != 0 ?
			((XSSFSheetImpl)sheet).shiftRowsRange(tRow, bRow, nRow, lCol, rCol, true, false, true, false, Range.FORMAT_NONE):
			((XSSFSheetImpl)sheet).shiftBothRange(tRow, bRow, nRow, lCol, rCol, nCol, true); //nCol != 0 && nRow != 0
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		shiftFormulas(all, sheet, tRow, bRow, nRow, lCol, rCol, nCol);
		all.add(new AreaRefImpl(tRow, lCol, bRow, rCol, refSheet)); //original selection
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		int ntRow = Math.max(0, tRow + nRow);
		int nlCol = Math.max(0, lCol + nCol);
		int nbRow = Math.min(maxrow, bRow + nRow);
		int nrCol = Math.min(maxcol, rCol + nCol);
		if (ntRow <= nbRow && nlCol <= nrCol)
		all.add(new AreaRefImpl(ntRow, nlCol, nbRow, nrCol, refSheet));
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	
	private static List<MergeChange> prepareChangeMerges(RefSheet sheet, List<CellRangeAddress[]> shiftedRanges) {
		final List<MergeChange> changeMerges = new ArrayList<MergeChange>();
		for(CellRangeAddress[] rngs : shiftedRanges) {
			final CellRangeAddress rng = rngs[1];
			Ref ref = null;
			if (rng != null) {
				int tRow = rng.getFirstRow();
				int lCol = rng.getFirstColumn();
				int bRow = rng.getLastRow();
				int rCol = rng.getLastColumn();
				ref = new AreaRefImpl(tRow, lCol, bRow, rCol, sheet);
			}
			final CellRangeAddress orgRng = rngs[0];
			int tRow = orgRng.getFirstRow();
			int lCol = orgRng.getFirstColumn();
			int bRow = orgRng.getLastRow();
			int rCol = orgRng.getLastColumn();
			final Ref orgRef = new AreaRefImpl(tRow, lCol, bRow, rCol, sheet);
			changeMerges.add(new MergeChange(orgRef, ref));
		}
		return changeMerges;
	}
	
	private static void shiftFormulas(Set<Ref> all, Sheet sheet, int startRow, int endRow, int nRow, int startCol, int endCol, int nCol) {
		final int moveSheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
        final PtgShifter shifter97 = new PtgShifter(moveSheetIndex, startRow, endRow, nRow, startCol, endCol, nCol, SpreadsheetVersion.EXCEL97);
        final PtgShifter shifter2007 = new PtgShifter(moveSheetIndex, startRow, endRow, nRow, startCol, endCol, nCol, SpreadsheetVersion.EXCEL2007);
		for (Ref ref : all) {
			final int tRow = ref.getTopRow();
			final int lCol = ref.getLeftCol();
			final Sheet srcSheet = getSheet(sheet, ref.getOwnerSheet());
			final Book srcBook = (Book) srcSheet.getWorkbook();
			final Cell srcCell = getCell(srcSheet, tRow, lCol);
			if (srcCell != null && srcCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				final int sheetIndex = srcBook.getSheetIndex(srcSheet);
				if (srcBook.getSpreadsheetVersion() == SpreadsheetVersion.EXCEL97) {
					shiftHSSFFormulas(shifter97, sheetIndex, srcCell);
				} else {
					shiftXSSFFormulas(shifter2007, sheetIndex, (XSSFCell)srcCell, (XSSFWorkbook)srcBook);
				}
			}
		}
	}
	
	private static void shiftHSSFFormulas(PtgShifter shifter, int sheetIndex, Cell cell) {
		CellValueRecordInterface vr = new HSSFCellHelper((HSSFCell)cell).getCellValueRecord();
		if (!(vr instanceof FormulaRecordAggregate)) {
			throw new IllegalArgumentException("Not a formula cell");
		}
		FormulaRecordAggregate fra = (FormulaRecordAggregate) vr;
		Ptg[] ptgs = fra.getFormulaTokens();//getFormulaRecord().getParsedExpression();
		if (shifter.adjustFormula(ptgs, sheetIndex)) {
			fra.setParsedExpression(ptgs);
		}
	}
	
	private static void shiftXSSFFormulas(PtgShifter shifter, int sheetIndex, XSSFCell cell, XSSFWorkbook book) {
        CTCell ctCell = cell.getCTCell();
        if (ctCell.isSetF()) {
            CTCellFormula f = ctCell.getF();
            String formula = f.getStringValue();
            if (formula.length() > 0) {
                String shiftedFormula = shiftXSSFFormula(book, sheetIndex, formula, shifter);
                if (shiftedFormula != null) {
                    f.setStringValue(shiftedFormula);
                }
            }

            if (f.isSetRef()) { //Range of cells which the formula applies to.
                String ref = f.getRef();
                String shiftedRef = shiftXSSFFormula(book, sheetIndex, ref, shifter);
                if (shiftedRef != null) f.setRef(shiftedRef);
            }
        }
	}

	private static String shiftXSSFFormula(XSSFWorkbook book, int sheetIndex, String formula, PtgShifter shifter) {
        XSSFEvaluationWorkbook fpb = XSSFEvaluationWorkbook.create(book);
        Ptg[] ptgs = FormulaParser.parse(formula, fpb, FormulaType.CELL, sheetIndex);
        String shiftedFmla = null;
        if (shifter.adjustFormula(ptgs, sheetIndex)) {
            shiftedFmla = FormulaRenderer.toFormulaString(fpb, ptgs);
        }
        return shiftedFmla;
	}
	
	public static ChangeInfo unMerge(Sheet sheet, int tRow, int lCol, int bRow, int rCol) {
		final RefSheet refSheet = BookHelper.getRefSheet((Book)sheet.getWorkbook(), sheet);
		final List<MergeChange> changes = new ArrayList<MergeChange>(); 
		for(int j = sheet.getNumMergedRegions() - 1; j >= 0; --j) {
        	final CellRangeAddress merged = sheet.getMergedRegion(j);
        	
        	final int firstCol = merged.getFirstColumn();
        	final int lastCol = merged.getLastColumn();
        	final int firstRow = merged.getFirstRow();
        	final int lastRow = merged.getLastRow();
        	if (firstCol >= lCol && lastCol <= rCol 
        		&& firstRow >= tRow && lastRow <= bRow) { //total cover
				changes.add(new MergeChange(new AreaRefImpl(firstRow, firstCol, lastRow, lastCol, refSheet), null));
				sheet.removeMergedRegion(j);
        	}
		}
		return new ChangeInfo(null, null, changes);
	}
	
	/**
	 * Merge the specified range per the given tRow, lCol, bRow, rCol.
	 * 
	 * @param sheet sheet
	 * @param tRow top row
	 * @param lCol left column
	 * @param bRow bottom row
	 * @param rCol right column
	 * @param across merge across each row.
	 * @return {@link Ref} array where the affected formula cell references in index 1 and to be evaluated formula cell references in index 0.
	 */
	@SuppressWarnings("unchecked")
	public static ChangeInfo merge(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean across) {
		if (across) {
			final Set<Ref> toEval = new HashSet<Ref>();
			final Set<Ref> affected = new HashSet<Ref>();
			final List<MergeChange> changes = new ArrayList<MergeChange>();
			for(int r = tRow; r <= bRow; ++r) {
				final ChangeInfo info = merge0(sheet, tRow, lCol, bRow, rCol);
				changes.addAll(info.getMergeChanges());
				toEval.addAll(info.getToEval());
				affected.addAll(info.getAffected());
			}
			return new ChangeInfo(toEval, affected, changes);
		} else {
			return merge0(sheet, tRow, lCol, bRow, rCol);
		}
	}
	
	private static ChangeInfo merge0(Sheet sheet, int tRow, int lCol, int bRow, int rCol) {
		final List<MergeChange> changes = new ArrayList<MergeChange>();
		final Set<Ref> all = new HashSet<Ref>();
		final Set<Ref> last = new HashSet<Ref>();
		//find the left most non-blank cell.
		Cell target = null;
		for(int r = tRow; target == null && r <= bRow; ++r) {
			for(int c = lCol; c <= rCol; ++c) {
				final Cell cell = BookHelper.getCell(sheet, r, c);
				if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK) {
					target = cell;
					break;
				}
			}
		}
		
		CellStyle style = null;
		if (target != null) { //found the target
			final int tgtRow = target.getRowIndex();
			final int tgtCol = target.getColumnIndex();
			final int nRow = tRow - tgtRow;
			final int nCol = lCol - tgtCol;
			if (nRow != 0 || nCol != 0) { //if target not the left-top one, move to left-top
				final ChangeInfo info = BookHelper.moveRange(sheet, tgtRow, tgtCol, tgtRow, tgtCol, nRow, nCol);
				if (info != null) {
					changes.addAll(info.getMergeChanges());
					last.addAll(info.getToEval());
					all.addAll(info.getAffected());
				}
			}
			final CellStyle source = target.getCellStyle();
			style = source.getIndex() == 0 ? null : sheet.getWorkbook().createCellStyle();
			if (style != null) {
				style.cloneStyleFrom(source);
				style.setBorderLeft(CellStyle.BORDER_NONE);
				style.setBorderTop(CellStyle.BORDER_NONE);
				style.setBorderRight(CellStyle.BORDER_NONE);
				style.setBorderBottom(CellStyle.BORDER_NONE);
				target.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
			}
			//1st row (exclude 1st cell)
			for (int c = lCol + 1; c <= rCol; ++c) {
				final Cell cell = getOrCreateCell(sheet, tRow, c);
				cell.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
				final Set<Ref>[] refs = BookHelper.setCellValue(cell, (RichTextString) null);
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
			//2nd row and after
			for(int r = tRow+1; r <= bRow; ++r) {
				for(int c = lCol; c <= rCol; ++c) {
					final Cell cell = getOrCreateCell(sheet, r, c);
					cell.setCellStyle(style); //set all cell in the merged range to CellStyle of the target minus border
					final Set<Ref>[] refs = BookHelper.setCellValue(cell, (RichTextString) null);
					if (refs != null) {
						last.addAll(refs[0]);
						all.addAll(refs[1]);
					}
				}
			}
		}
		
		sheet.addMergedRegion(new CellRangeAddress(tRow, bRow, lCol, rCol));
		final Ref mergeArea = new AreaRefImpl(tRow, lCol, bRow, rCol, BookHelper.getRefSheet((Book)sheet.getWorkbook(), sheet)); 
		all.add(mergeArea);
		changes.add(new MergeChange(null, mergeArea));
		
		return new ChangeInfo(last, all, changes);
	}
	
	@SuppressWarnings("unchecked")
	public static Set<Ref>[] sort(Sheet sheet, int tRow, int lCol, int bRow, int rCol, 
			Ref key1, boolean desc1, Ref key2, int type, boolean desc2, Ref key3, boolean desc3, int header, int orderCustom,
			boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3) {
		//TODO type not yet imiplmented(Sort label/Sort value, for PivotTable)
		//TODO orderCustom is not implemented yet
		if (header == BookHelper.SORT_HEADER_YES) {
			if (sortByRows) {
				++lCol;
			} else {
				++tRow;
			}
		}
		if (tRow > bRow || lCol > rCol) { //nothing to sort!
			return (Set<Ref>[]) new HashSet[0];
		}
		int keyCount = 0;
		if (key1 != null) {
			++keyCount;
			if (key2 != null) {
				++keyCount;
				if (key3 != null) {
					++keyCount;
				}
			}
		}
		if (keyCount == 0) {
			throw new UiException("Must specify at least the key1");
		}
		final int[] dataOptions = new int[keyCount];
		final boolean[] descs = new boolean[keyCount];
		final int[] keyIndexes = new int[keyCount];
		keyIndexes[0] = rangeToIndex(key1, sortByRows);
		descs[0] = desc1;
		dataOptions[0] = dataOption1;
		if (keyCount > 1) {
			keyIndexes[1] = rangeToIndex(key2, sortByRows);
			descs[1] = desc2;
			dataOptions[1] = dataOption2;
		}
		if (keyCount > 2) {
			keyIndexes[2] = rangeToIndex(key3, sortByRows);
			descs[2] = desc3;
			dataOptions[2] = dataOption3;
		}
		validateKeyIndexes(keyIndexes, tRow, lCol, bRow, rCol, sortByRows);
		
		final List<SortKey> sortKeys = new ArrayList<SortKey>(sortByRows ? rCol - lCol + 1 : bRow - tRow + 1);
		final int begRow = Math.max(tRow, sheet.getFirstRowNum());
		final int endRow = Math.min(bRow, sheet.getLastRowNum());
		if (sortByRows) {
			int begCol = ((Book)sheet.getWorkbook()).getSpreadsheetVersion().getLastColumnIndex();
			int endCol = 0;
			//locate begCol/endCol of the sheet
			for (int rowNum = begRow; rowNum <= endRow; ++rowNum) {
				final Row row = sheet.getRow(rowNum);
				if (row != null) {
					begCol = Math.min(begCol, row.getFirstCellNum());
					endCol = Math.max(begCol, row.getLastCellNum() - 1);
				}
			}
			begCol = Math.max(lCol, begCol);
			endCol = Math.min(rCol, endCol);
			for (int colnum = begCol; colnum <= endCol; ++colnum) {
				final Object[] values = new Object[keyCount];
				for(int j = 0; j < keyCount; ++j) {
					final Row row = sheet.getRow(keyIndexes[j]);
					final Cell cell = row != null ? row.getCell(colnum, Row.RETURN_BLANK_AS_NULL) : null;
					final Object val = getCellObject(cell, dataOptions[j]);
					values[j] = val;
				}
				final SortKey sortKey = new SortKey(colnum, values);
				sortKeys.add(sortKey);
			}
			if (!sortKeys.isEmpty()) {
				final Comparator<SortKey> keyComparator = new KeyComparator(descs, matchCase, sortMethod, type);
				Collections.sort(sortKeys, keyComparator);
				return BookHelper.assignColumns(sheet, sortKeys, begRow, lCol, endRow, rCol);
			} else {
				return null;
			}
		} else { //sortByColumn, default case
			for (int rownum = begRow; rownum <= endRow; ++rownum) {
				final Row row = sheet.getRow(rownum);
				if (row == null) {
					continue; //nothing to sort
				}
				final Object[] values = new Object[keyCount];
				for(int j = 0; j < keyCount; ++j) {
					final Cell cell = row.getCell(keyIndexes[j], Row.RETURN_BLANK_AS_NULL);
					final Object val = getCellObject(cell, dataOptions[j]);
					values[j] = val;
				}
				final SortKey sortKey = new SortKey(rownum, values);
				sortKeys.add(sortKey);
			}
			if (!sortKeys.isEmpty()) {
				final Comparator<SortKey> keyComparator = new KeyComparator(descs, matchCase, sortMethod, type);
				Collections.sort(sortKeys, keyComparator);
				return BookHelper.assignRows(sheet, sortKeys, tRow, lCol, bRow, rCol);
			} else {
				return null;
			}
		}
	}
	
	private static Object getCellObject(Cell cell, int dataOption) {
		Object val = cell != null ? BookHelper.getCellObject(cell) : null;
		if (val instanceof RichTextString && dataOption == BookHelper.SORT_TEXT_AS_NUMBERS) {
			try {
				val = new Double((String)((RichTextString)val).getString());
			} catch(NumberFormatException ex) {
				val = new Double(0);//ignore
			}
		}
		return val;
	}
	@SuppressWarnings("unchecked")
	private static Set<Ref>[] assignColumns(Sheet sheet, List<SortKey> sortKeys, int tRow, int lCol, int bRow, int rCol) {
		final int cellCount = bRow - tRow + 1;
		final Map<Integer, List<Cell>> newCols = new HashMap<Integer, List<Cell>>();  
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		int j = 0;
		for(final Iterator<SortKey> it = sortKeys.iterator(); it.hasNext();++j) {
			final SortKey sortKey = it.next();
			final int oldColNum = sortKey.getIndex();
			final int newColNum = lCol + j;
			it.remove();
			if (oldColNum == newColNum) { //no move needed, skip it
				continue;
			}
			//remove cells from the old column of the Range
			final List<Cell> cells = new ArrayList<Cell>(cellCount);
			for(int k = tRow; k <= bRow; ++k) {
				final Cell cell = BookHelper.getCell(sheet, k, oldColNum);
				if (cell != null) {
					cells.add(cell);
					final Set<Ref>[] refs = BookHelper.removeCell(cell, false);
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
			if (!cells.isEmpty()) {
				newCols.put(new Integer(newColNum), cells);
			}
		}
		
		//move cells
		for(Entry<Integer, List<Cell>> entry : newCols.entrySet()) {
			final int colNum = entry.getKey().intValue();
			final List<Cell> cells = entry.getValue();
			for(Cell cell : cells) {
				final int rowNum = cell.getRowIndex();
				final Set<Ref>[] refs = BookHelper.copyCell(cell, sheet, rowNum, colNum, Range.PASTE_ALL, Range.PASTEOP_NONE);
				last.addAll(refs[0]);
				all.addAll(refs[1]);
			}
		}
		return (Set<Ref>[]) new Set[] {last, all};
	}

	@SuppressWarnings("unchecked")
	private static Set<Ref>[] assignRows(Sheet sheet, List<SortKey> sortKeys, int tRow, int lCol, int bRow, int rCol) {
		final int cellCount = rCol - lCol + 1;
		final Map<Integer, List<Cell>> newRows = new HashMap<Integer, List<Cell>>();  
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		int j = 0;
		for(final Iterator<SortKey> it = sortKeys.iterator(); it.hasNext();++j) {
			final SortKey sortKey = it.next();
			final int oldRowNum = sortKey.getIndex();
			final Row row = sheet.getRow(oldRowNum); 
			final int newRowNum = tRow + j;
			it.remove();
			if (oldRowNum == newRowNum) { //no move needed, skip it
				continue;
			}
			//remove cells from the old row of the Range
			final List<Cell> cells = new ArrayList<Cell>(cellCount);
			final int begCol = Math.max(lCol, row.getFirstCellNum());
			final int endCol = Math.min(rCol, row.getLastCellNum() - 1);
			for(int k = begCol; k <= endCol; ++k) {
				final Cell cell = row.getCell(k);
				if (cell != null) {
					cells.add(cell);
					final Set<Ref>[] refs = BookHelper.removeCell(cell, false);
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
			if (!cells.isEmpty()) {
				newRows.put(new Integer(newRowNum), cells);
			}
		}
		
		//move cells
		for(Entry<Integer, List<Cell>> entry : newRows.entrySet()) {
			final int rowNum = entry.getKey().intValue();
			final List<Cell> cells = entry.getValue();
			for(Cell cell : cells) {
				final int colNum = cell.getColumnIndex();
				final Set<Ref>[] refs = BookHelper.copyCell(cell, sheet, rowNum, colNum, Range.PASTE_ALL, Range.PASTEOP_NONE);
				last.addAll(refs[0]);
				all.addAll(refs[1]);
			}
		}
		return (Set<Ref>[]) new Set[] {last, all};
	}
	private static Object getCellObject(Cell cell) {
		if (cell == null) {
			return "";
		}
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_FORMULA) {
			final Book book = (Book)cell.getSheet().getWorkbook();
			final CellValue cv = BookHelper.evaluate(book, cell);
			return BookHelper.getValueByCellValue(cv);
		} else {
			return BookHelper.getCellValue(cell);
		}
	}
	private static int rangeToIndex(Ref range, boolean sortByRows) {
		return sortByRows ? range.getTopRow() : range.getLeftCol();
	}
	private static void validateKeyIndexes(int[] keyIndexes, int tRow, int lCol, int bRow, int rCol, boolean sortByRows) {
		if (!sortByRows) {
			for(int j = keyIndexes.length - 1; j >= 0; --j) {
				final int keyIndex = keyIndexes[j]; 
				if (keyIndex < lCol || keyIndex > rCol) {
					throw new UiException("The given key is out of the sorting range: "+keyIndex);
				}
			}
		} else {
			for(int j = keyIndexes.length - 1; j >= 0; --j) {
				final int keyIndex = keyIndexes[j]; 
				if (keyIndex < tRow || keyIndex > bRow) {
					throw new UiException("The given key is out of the sorting range: "+keyIndex);
				}
			}
		}
	}
	
	public static class SortKey {
		final private int _index; //original row/column index
		final private Object[] _values;
		public SortKey(int index, Object[] values) {
			this._index = index;
			this._values = values;
		}
		public int getIndex() {
			return _index;
		}
		public Object[] getValues() {
			return _values;
		}
	}
	
	private static class KeyComparator implements Comparator<SortKey> {
		final private boolean[] _descs;
		final private boolean _matchCase;
		final private int _sortMethod; //TODO byNumberOfStorks, byPinyYin
		final private int _type; //TODO PivotTable only: byLabel, byValue
		
		public KeyComparator(boolean[] descs, boolean matchCase, int sortMethod, int type) {
			_descs = descs;
			_matchCase = matchCase;
			_sortMethod = sortMethod;
			_type = type;
		}
		@Override
		public int compare(SortKey o1, SortKey o2) {
			final Object[] values1 = o1.getValues();
			final Object[] values2 = o2.getValues();
			return compare(values1, values2);
		}

		private int compare(Object[] values1, Object[] values2) {
			final int len = values1.length;
			for(int j = 0; j < len; ++j) {
				int p = compareValue(values1[j], values2[j], _descs[j]);
				if (p != 0) {
					return p;
				}
			}
			return 0;
		}
		//1. null is always sorted at the end
		//2. Error(Byte) > Boolean > String > Double
		private int compareValue(Object val1, Object val2, boolean desc) {
			if (val1 == val2) {
				return 0;
			}
			final int order1 = val1 instanceof Byte ? 4 : val1 instanceof Boolean ? 3 : val1 instanceof RichTextString ? 2 : val1 instanceof Number ? 1 : desc ? 0 : 5;
			final int order2 = val2 instanceof Byte ? 4 : val2 instanceof Boolean ? 3 : val2 instanceof RichTextString ? 2 : val2 instanceof Number ? 1 : desc ? 0 : 5;
			int ret = 0;
			if (order1 != order2) {
				ret = order1 - order2;
			} else { //order1 == order2
				switch(order1) {
				case 4: //error, no order among different errors
					ret = 0;
					break;
				case 3: //Boolean
					ret = ((Boolean)val1).compareTo((Boolean)val2);
					break;
				case 2: //RichTextString
					ret = compareString(((RichTextString)val1).getString(), ((RichTextString)val2).getString());
					break;
				case 1: //Double
					ret =((Double)val1).compareTo((Double)val2);
					break;
				default:
					throw new UiException("Unknown value type: "+val1.getClass());
				}
			}
			return desc ? -ret : ret;
		}
		private int compareString(String s1, String s2) {
			return _matchCase ? compareString0(s1, s2) : s1.compareToIgnoreCase(s2);
		}
		//bug 59 Sort with case sensitive should be in special spreadsheet order
		private int compareString0(String s1, String s2) {
			final int len1 = s1.length();
			final int len2 = s2.length();
			final int len = len1 > len2 ? len2 : len1;
			for (int j = 0; j < len; ++j) {
				final int ret = compareChar(s1.charAt(j), s2.charAt(j));
				if ( ret != 0) {
					return ret;
				}
			}
			return len1 - len2;
		}
		private int compareChar(char ch1, char ch2) {
			final char uch1 = Character.toUpperCase(ch1);
			final char uch2 = Character.toUpperCase(ch2);
			return uch1 == uch2 ? 
					(ch2 - ch1) : //yes, a < A
					(uch1 - uch2); //yes, a < b, a < B, A < b, and A < B
		}
	}
	 
	public static Font getOrCreateFont(Book book, short boldWeight, Color color, short fontHeight, java.lang.String name, 
			boolean italic, boolean strikeout, short typeOffset, byte underline) {
		Font font = book.findFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
		if (font == null) {
			font = book.createFont();
			font.setBoldweight(boldWeight);
			BookHelper.setFontColor(font, color);
			font.setFontHeight(fontHeight);
			font.setFontName(name);
			font.setItalic(italic);
			font.setStrikeout(strikeout);
			font.setTypeOffset(typeOffset);
			font.setUnderline(underline);
		}
		return font;
	}
	
	public static Font getFont(Cell cell) {
		final CellStyle style = cell.getCellStyle();
		final short fontIdx = style.getFontIndex();
		return cell.getSheet().getWorkbook().getFontAt(fontIdx);
	}
	
	private final static Map<BorderStyle, Short> _BORDER_STYLE_INDEX = new HashMap<BorderStyle, Short>();
	static {
		_BORDER_STYLE_INDEX.put(BorderStyle.DASH_DOT, CellStyle.BORDER_DASH_DOT);
		_BORDER_STYLE_INDEX.put(BorderStyle.DASH_DOT_DOT, CellStyle.BORDER_DASH_DOT_DOT);
		_BORDER_STYLE_INDEX.put(BorderStyle.DASHED, CellStyle.BORDER_DASHED);
		_BORDER_STYLE_INDEX.put(BorderStyle.DOTTED, CellStyle.BORDER_DOTTED);
		_BORDER_STYLE_INDEX.put(BorderStyle.DOUBLE, CellStyle.BORDER_DOUBLE);
		_BORDER_STYLE_INDEX.put(BorderStyle.HAIR, CellStyle.BORDER_HAIR);
		_BORDER_STYLE_INDEX.put(BorderStyle.MEDIUM, CellStyle.BORDER_MEDIUM);
		_BORDER_STYLE_INDEX.put(BorderStyle.MEDIUM_DASH_DOT, CellStyle.BORDER_MEDIUM_DASH_DOT);
		_BORDER_STYLE_INDEX.put(BorderStyle.MEDIUM_DASH_DOT_DOT, CellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
		_BORDER_STYLE_INDEX.put(BorderStyle.MEDIUM_DASHED, CellStyle.BORDER_MEDIUM_DASHED);
		_BORDER_STYLE_INDEX.put(BorderStyle.NONE, CellStyle.BORDER_NONE);
		_BORDER_STYLE_INDEX.put(BorderStyle.SLANTED_DASH_DOT, CellStyle.BORDER_SLANTED_DASH_DOT);
		_BORDER_STYLE_INDEX.put(BorderStyle.THICK, CellStyle.BORDER_THICK);
		_BORDER_STYLE_INDEX.put(BorderStyle.THIN, CellStyle.BORDER_THIN);
	}
	private static short getBorderStyleIndex(BorderStyle lineStyle) {
		return _BORDER_STYLE_INDEX.get(lineStyle).shortValue();
	}
	
	public static Set<Ref> setBorders(Sheet sheet, int tRow, int lCol, int bRow, int rCol, short borderIndex, BorderStyle lineStyle, String color) {
		Set<Ref> all = null;
		if ((borderIndex & BookHelper.BORDER_INSIDE) != 0) {
			all = setBordersInside(sheet, tRow, lCol, bRow, rCol, borderIndex, lineStyle, color);
		}
		if ((borderIndex & BookHelper.BORDER_DIAGONAL) != 0) {
			final Set<Ref> refs = setBordersDiagonal(sheet, tRow, lCol, bRow, rCol, borderIndex, lineStyle, color); 
			if (all == null) {
				all = refs;
			}
		}
		if ((borderIndex & BookHelper.BORDER_OUTLINE) != 0) {
			final Set<Ref> refs = setBordersOutline(sheet, tRow, lCol, bRow, rCol, borderIndex, lineStyle, color);
			if (all == null) {
				all = refs;
			}
		}
		return all;
	}
	
	private static Set<Ref> setBordersOutline(Sheet sheet, int tRow, int lCol, int bRow, int rCol, short borderIndex, BorderStyle lineStyle, String color) {
		final Book book = (Book) sheet.getWorkbook();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
		final Color bsColor = BookHelper.HTMLToColor(book, color);
		final short bsLineStyle = getBorderStyleIndex(lineStyle);
		Set<Ref> all = new HashSet<Ref>();
		
		final int lb = BookHelper.BORDER_EDGE_LEFT;// leftBorder
		final int tb = BookHelper.BORDER_EDGE_TOP;// topBorder
		final int rb = BookHelper.BORDER_EDGE_RIGHT;// rightBorder
		final int bb = BookHelper.BORDER_EDGE_BOTTOM;// bottomBorder
		
		//top line
		if ((borderIndex & tb) != 0) {
			for (int col = lCol; col <= rCol; ++col) {
				Styles.setBorder(sheet, tRow, col, bsColor, bsLineStyle, tb);
			}
			all.add(new AreaRefImpl(tRow, lCol, tRow, rCol, refSheet));
			final int tRow0 = tRow - 1;
			if (tRow0 >= 0) {
				for (int col = lCol; col <= rCol; ++col) {
					Styles.setBorder(sheet, tRow0, col, bsColor, bsLineStyle, bb);
				}
				all.add(new AreaRefImpl(tRow0, lCol, tRow0, rCol, refSheet));
			}
		}
		
		//bottom row
		if ((borderIndex & bb) != 0) {
			for (int col = lCol; col <= rCol; ++col) {
				Styles.setBorder(sheet, bRow, col, bsColor, bsLineStyle, bb);
			}
			all.add(new AreaRefImpl(bRow, lCol, bRow, rCol, refSheet));
			final int bRow0 = bRow + 1;
			if (bRow0 <= maxrow) {
				for (int col = lCol; col <= rCol; ++col) {
					Styles.setBorder(sheet, bRow0, col, bsColor, bsLineStyle, tb);
				}
				all.add(new AreaRefImpl(bRow0, lCol, bRow0, rCol, refSheet));
			}
		}
		
		//left column
		if ((borderIndex & lb) != 0) {
			for (int row = tRow; row <= bRow; row++) {
				Styles.setBorder(sheet, row, lCol, bsColor, bsLineStyle, lb);
			}
			all.add(new AreaRefImpl(tRow, lCol, bRow, lCol, refSheet));
			final int lCol0 = lCol - 1;
			if (lCol0 >= 0) {
				for (int row = tRow; row <= bRow; row++) {
					Styles.setBorder(sheet, row, lCol0, bsColor, bsLineStyle, rb);
				}
				all.add(new AreaRefImpl(tRow, lCol0, bRow, lCol0, refSheet));
			}
		}

		//right column
		if ((borderIndex & rb) != 0) {
			for (int row = tRow; row <= bRow; row++) {
				Styles.setBorder(sheet, row, rCol, bsColor, bsLineStyle, rb);
			}
			all.add(new AreaRefImpl(tRow, rCol, bRow, rCol, refSheet));
			final int rCol0 = rCol + 1;
			if (rCol0 <= maxcol) {
				for (int row = tRow; row <= bRow; row++) {
					Styles.setBorder(sheet, row, rCol0, bsColor, bsLineStyle, lb);
				}
				all.add(new AreaRefImpl(tRow, rCol0, bRow, rCol0, refSheet));
			}
		}

		return all;
	}
	
	private static Set<Ref> setBordersInside(Sheet sheet, int tRow, int lCol, int bRow, int rCol, short borderIndex, BorderStyle lineStyle, String color) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
		final Color bsColor = BookHelper.HTMLToColor(book, color);
		final short bsLineStyle = getBorderStyleIndex(lineStyle);
		Set<Ref> all = new HashSet<Ref>();
		
		final int lb = BookHelper.BORDER_EDGE_LEFT;// leftBorder
		final int tb = BookHelper.BORDER_EDGE_TOP;// topBorder
		final int rb = BookHelper.BORDER_EDGE_RIGHT;// rightBorder
		final int bb = BookHelper.BORDER_EDGE_BOTTOM;// bottomBorder
		final int hb = BookHelper.BORDER_INSIDE_HORIZONTAL;// horizontal Border
		final int vb = BookHelper.BORDER_INSIDE_VERTICAL;// vertical Border
		
		//vertical borders
		if ((borderIndex & vb) != 0) {
			for (int row = tRow; row <= bRow; ++row) {
				for (int col = lCol; col < rCol; ++col) {
					Styles.setBorder(sheet, tRow, col, bsColor, bsLineStyle, rb);
				}
			}
			for (int row = tRow; row <= bRow; ++row) {
				for (int col = lCol+1; col <= rCol; ++col) {
					Styles.setBorder(sheet, tRow, col, bsColor, bsLineStyle, lb);
				}
			}
		}
		
		//horizontal borders
		if ((borderIndex & hb) != 0) {
			for (int row = tRow; row < bRow; ++row) {
				for (int col = lCol; col <= rCol; ++col) {
					Styles.setBorder(sheet, tRow, col, bsColor, bsLineStyle, bb);
				}
			}
			for (int row = tRow + 1; row <= bRow; ++row) {
				for (int col = lCol; col < rCol; ++col) {
					Styles.setBorder(sheet, tRow, col, bsColor, bsLineStyle, tb);
				}
			}
		}
		
		all.add(new AreaRefImpl(tRow, lCol, bRow, rCol, refSheet));
		return all;
	}

	private static Set<Ref> setBordersDiagonal(Sheet sheet, int tRow, int lCol, int bRow, int rCol, short borderIndex, BorderStyle lineStyle, String color) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
		final Color bsColor = BookHelper.HTMLToColor(book, color);
		final short bsLineStyle = getBorderStyleIndex(lineStyle);
		Set<Ref> all = new HashSet<Ref>();
		
		final int db = BookHelper.BORDER_DIAGONAL_DOWN;// diagonal down Border
		final int ub = BookHelper.BORDER_DIAGONAL_UP;// diagonal up Border
		
		//diagonal down borders
		if ((borderIndex & db) != 0) {
			for (int row = tRow; row <= bRow; ++row) {
				for (int col = lCol; col <= rCol; ++col) {
					Styles.setBorder(sheet, tRow, col, bsColor, bsLineStyle, db);
				}
			}
		}
		
		//diagonal up borders
		if ((borderIndex & ub) != 0) {
			for (int row = tRow; row <= bRow; ++row) {
				for (int col = lCol; col <= rCol; ++col) {
					Styles.setBorder(sheet, tRow, col, bsColor, bsLineStyle, ub);
				}
			}
		}
		
		all.add(new AreaRefImpl(tRow, lCol, bRow, rCol, refSheet));
		return all;
	}
	
	public static Set<Ref> setColumnWidth(Sheet sheet, int lCol, int rCol, int char256) {
		final Book book = (Book) sheet.getWorkbook();
		final int maxrow = book.getSpreadsheetVersion().getLastRowIndex();
		final RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
		final Set<Ref> all = new HashSet<Ref>();
		for (int col = lCol; col <= rCol; ++col) {
			final int orgChar256 = sheet.getColumnWidth(col);
			if (char256 != orgChar256) {
				sheet.setColumnWidth(col, char256);
				all.add(new AreaRefImpl(0, col, maxrow, col, refSheet));
			}
		}
		return all;
	}
	
	public static Set<Ref> setRowHeight(Sheet sheet, int tRow, int bRow, short twips) {
		final Book book = (Book) sheet.getWorkbook();
		final int maxcol = book.getSpreadsheetVersion().getLastColumnIndex();
		final RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
		final Set<Ref> all = new HashSet<Ref>();
		for (int row = tRow; row <= bRow; ++row) {
			final int orgTwips = sheet.getColumnWidth(row);
			if (twips != orgTwips) {
				BookHelper.setRowHeight(sheet, row, twips);
				all.add(new AreaRefImpl(row, 0, row, maxcol, refSheet));
			}
		}
		return all;
	}
	
	public static void setRowHeight(Sheet sheet, int row, short twips) {
		final Row rowx = BookHelper.getOrCreateRow(sheet, row);
		rowx.setHeight(twips);
	}
	
	public static short getRowHeight(Sheet sheet, int row) {
		final Row rowx = sheet.getRow(row);
		return rowx != null ? getRowHeight(rowx) : sheet.getDefaultRowHeight();
	}
	
	private static short getRowHeight(Row row) {
		final short h = row.getHeight();
		return h == 0xFF ? row.getSheet().getDefaultRowHeight() : h;
	}
	
	public static Set<Ref> setCellStyle(Sheet sheet, int tRow, int lCol, int bRow, int rCol, CellStyle style) {
		for(int r = tRow; r <= bRow; ++r) {
			for (int c = lCol; c <= rCol; ++c) {
				final Cell cell = BookHelper.getOrCreateCell(sheet, r, c);
				cell.setCellStyle(style);
			}
		}
		final Set<Ref> all = new HashSet<Ref>(1);
		final RefSheet refSheet = BookHelper.getRefSheet((Book)sheet.getWorkbook(), sheet);
		all.add(new AreaRefImpl(tRow, lCol, bRow, rCol, refSheet));
		
		return all; 
	}
	
	//inner fill direction for #fill
	private static final int FILL_INVALID = 0;
	private static final int FILL_NONE = 1; //no way to fill
	private static final int FILL_UP = 2;
	private static final int FILL_DOWN = 3;
	private static final int FILL_RIGHT = 4;
	private static final int FILL_LEFT = 5;
	
	//[0]:last [1]:all
	public static Set<Ref>[] fill(Sheet sheet, Ref srcRef, Ref dstRef, int fillType) {
//TODO, FILL_DEFAULT shall check the contents of the source cell, now we default to FILL_COPY
if (fillType == FILL_DEFAULT) {
	fillType = FILL_COPY;
}
		final int fillDir = BookHelper.getFillDirection(sheet, srcRef, dstRef);
		if (fillDir == BookHelper.FILL_NONE) { //nothing to fill up, just return
			return null;
		}
		switch(fillDir) {
		case FILL_UP:
			return fillUp(sheet, srcRef, dstRef, fillType);
		case FILL_DOWN:
			return fillDown(sheet, srcRef, dstRef, fillType);
		case FILL_RIGHT:
			return fillRight(sheet, srcRef, dstRef, fillType);
		case FILL_LEFT:
			return fillLeft(sheet, srcRef, dstRef, fillType);
		}
		//FILL_INVALID
		throw new UiException("Destination range must include source range and can be fill in one direction only"); 
	}
	
	@SuppressWarnings("unchecked")
	public static Set<Ref>[] fillDown(Sheet sheet, Ref srcRef, Ref dstRef, int fillType) {
		//TODO FILL_DEFAULT, FILL_DAYS, FILL_WEEKDAYS, FILL_MONTHS, FILL_YEARS, FILL_GROWTH_TREND, FILL_LINER_TREND, FILL_SERIES
		int pasteType = BookHelper.INNERPASTE_FILL_COPY;
		switch(fillType) {
		case FILL_COPY:
			pasteType = BookHelper.INNERPASTE_FILL_COPY;
			break;
		case FILL_FORMATS:
			pasteType = BookHelper.INNERPASTE_FILL_FORMATS;
			break;
		case FILL_VALUES:
			pasteType = BookHelper.INNERPASTE_FILL_VALUE;
			break;
		default:
			return null;
		}
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		final int rowCount = srcRef.getRowCount();
		final int srctRow = srcRef.getTopRow();
		final int srcbRow = srcRef.getBottomRow();
		final int srclCol = srcRef.getLeftCol();
		final int srcrCol = srcRef.getRightCol();
		
		final int dstbRow = dstRef.getBottomRow();
		for(int srcIndex = 0, r = srcbRow + 1; r <= dstbRow; ++r, ++srcIndex) {
			for(int c = srclCol; c <= srcrCol; ++c) {
				final int srcrow = srctRow + srcIndex % rowCount;
				final Cell srcCell = BookHelper.getCell(sheet, srcrow, c);
				final Set<Ref>[] refs = srcCell == null ? 
						BookHelper.removeCell(sheet, r, c) : BookHelper.copyCell(srcCell, sheet, r, c, pasteType, BookHelper.PASTEOP_NONE); 
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
		}
		final RefSheet refSheet = BookHelper.getRefSheet((Book)sheet.getWorkbook(), sheet);
		all.add(new AreaRefImpl(srcbRow + 1, srclCol, dstbRow, srcrCol, refSheet));
		return (Set<Ref>[]) new Set[] {last, all};
	}
	
	@SuppressWarnings("unchecked")
	public static Set<Ref>[] fillUp(Sheet sheet, Ref srcRef, Ref dstRef, int fillType) {
		//TODO FILL_DEFAULT, FILL_DAYS, FILL_WEEKDAYS, FILL_MONTHS, FILL_YEARS, FILL_GROWTH_TREND, FILL_LINER_TREND, FILL_SERIES
		int pasteType = BookHelper.INNERPASTE_FILL_COPY;
		switch(fillType) {
		case FILL_COPY:
			pasteType = BookHelper.INNERPASTE_FILL_COPY;
			break;
		case FILL_FORMATS:
			pasteType = BookHelper.INNERPASTE_FILL_FORMATS;
			break;
		case FILL_VALUES:
			pasteType = BookHelper.INNERPASTE_FILL_VALUE;
			break;
		default:
			return null;
		}
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		final int rowCount = srcRef.getRowCount();
		final int srctRow = srcRef.getTopRow();
		final int srcbRow = srcRef.getBottomRow();
		final int srclCol = srcRef.getLeftCol();
		final int srcrCol = srcRef.getRightCol();
		
		final int dsttRow = dstRef.getTopRow();
		for(int srcIndex = 0, r = srctRow - 1; r >= dsttRow; --r, ++srcIndex) {
			for(int c = srclCol; c <= srcrCol; ++c) {
				final int srcrow = srcbRow - srcIndex % rowCount;
				final Cell srcCell = BookHelper.getCell(sheet, srcrow, c);
				final Set<Ref>[] refs = srcCell == null ? 
						BookHelper.removeCell(sheet, r, c) : BookHelper.copyCell(srcCell, sheet, r, c, pasteType, BookHelper.PASTEOP_NONE); 
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
		}
		final RefSheet refSheet = BookHelper.getRefSheet((Book)sheet.getWorkbook(), sheet);
		all.add(new AreaRefImpl(dsttRow, srclCol, srctRow - 1, srcrCol, refSheet));
		return (Set<Ref>[]) new Set[] {last, all};
	}
	
	@SuppressWarnings("unchecked")
	public static Set<Ref>[] fillRight(Sheet sheet, Ref srcRef, Ref dstRef, int fillType) {
		//TODO FILL_DEFAULT, FILL_DAYS, FILL_WEEKDAYS, FILL_MONTHS, FILL_YEARS, FILL_GROWTH_TREND, FILL_LINER_TREND, FILL_SERIES
		int pasteType = BookHelper.INNERPASTE_FILL_COPY;
		switch(fillType) {
		case FILL_COPY:
			pasteType = BookHelper.INNERPASTE_FILL_COPY;
			break;
		case FILL_FORMATS:
			pasteType = BookHelper.INNERPASTE_FILL_FORMATS;
			break;
		case FILL_VALUES:
			pasteType = BookHelper.INNERPASTE_FILL_VALUE;
			break;
		default:
			return null;
		}
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		final int colCount = srcRef.getColumnCount();
		final int srclCol = srcRef.getLeftCol();
		final int srcrCol = srcRef.getRightCol();
		final int srctRow = srcRef.getTopRow();
		final int srcbRow = srcRef.getBottomRow();
		
		final int dstrCol = dstRef.getRightCol();
		for(int r = srctRow; r <= srcbRow; ++r) {
			for(int srcIndex = 0, c = srcrCol + 1; c <= dstrCol; ++c, ++srcIndex) {
				final int srccol = srclCol + srcIndex % colCount;
				final Cell srcCell = BookHelper.getCell(sheet, r, srccol);
				final Set<Ref>[] refs = srcCell == null ? 
						BookHelper.removeCell(sheet, r, c) : BookHelper.copyCell(srcCell, sheet, r, c, pasteType, BookHelper.PASTEOP_NONE); 
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
		}
		final RefSheet refSheet = BookHelper.getRefSheet((Book)sheet.getWorkbook(), sheet);
		all.add(new AreaRefImpl(srctRow, srcrCol + 1, srcbRow, dstrCol, refSheet));
		return (Set<Ref>[]) new Set[] {last, all};
	}
	
	@SuppressWarnings("unchecked")
	public static Set<Ref>[] fillLeft(Sheet sheet, Ref srcRef, Ref dstRef, int fillType) {
		//TODO FILL_DEFAULT, FILL_DAYS, FILL_WEEKDAYS, FILL_MONTHS, FILL_YEARS, FILL_GROWTH_TREND, FILL_LINER_TREND, FILL_SERIES
		int pasteType = BookHelper.INNERPASTE_FILL_COPY;
		switch(fillType) {
		case FILL_COPY:
			pasteType = BookHelper.INNERPASTE_FILL_COPY;
			break;
		case FILL_FORMATS:
			pasteType = BookHelper.INNERPASTE_FILL_FORMATS;
			break;
		case FILL_VALUES:
			pasteType = BookHelper.INNERPASTE_FILL_VALUE;
			break;
		default:
			return null;
		}
		final Set<Ref> last = new HashSet<Ref>();
		final Set<Ref> all = new HashSet<Ref>();
		final int colCount = srcRef.getColumnCount();
		final int srclCol = srcRef.getLeftCol();
		final int srcrCol = srcRef.getRightCol();
		final int srctRow = srcRef.getTopRow();
		final int srcbRow = srcRef.getBottomRow();
		
		final int dstlCol = dstRef.getLeftCol();
		for(int r = srctRow; r <= srcbRow; ++r) {
			for(int srcIndex = 0, c = srclCol - 1; c >= dstlCol; --c, ++srcIndex) {
				final int srccol = srcrCol - srcIndex % colCount;
				final Cell srcCell = BookHelper.getCell(sheet, r, srccol);
				final Set<Ref>[] refs = srcCell == null ? 
						BookHelper.removeCell(sheet, r, c) : BookHelper.copyCell(srcCell, sheet, r, c, pasteType, BookHelper.PASTEOP_NONE); 
				if (refs != null) {
					last.addAll(refs[0]);
					all.addAll(refs[1]);
				}
			}
		}
		final RefSheet refSheet = BookHelper.getRefSheet((Book)sheet.getWorkbook(), sheet);
		all.add(new AreaRefImpl(srctRow, dstlCol, srcbRow, srclCol - 1, refSheet));
		return (Set<Ref>[]) new Set[] {last, all};
	}
	
	private static int getFillDirection(Sheet sheet, Ref srcRef, Ref dstRef) {
		final Sheet dstSheet = BookHelper.getSheet(sheet, dstRef.getOwnerSheet());
		if (dstSheet.equals(sheet)) {
			final int dsttRow = dstRef.getTopRow();
			final int dstbRow = dstRef.getBottomRow();
			final int dstlCol = dstRef.getLeftCol();
			final int dstrCol = dstRef.getRightCol();
			
			final int srctRow = srcRef.getTopRow();
			final int srcbRow = srcRef.getBottomRow();
			final int srclCol = srcRef.getLeftCol();
			final int srcrCol = srcRef.getRightCol();
			
			//check fill direction
			if (srclCol == dstlCol && srcrCol == dstrCol) {
				if (dsttRow == srctRow) {
					if (dstbRow > srcbRow) { //fill down
						return FILL_DOWN;
					} else if (dstbRow == srcbRow) { //nothing to fill
						return FILL_NONE;
					}
				}
				if (dstbRow == srcbRow && dsttRow < srctRow) { //fill up
					return FILL_UP;
				}
			} else if (srctRow == dsttRow && srcbRow == dstbRow) {
				if (dstlCol == srclCol && dstrCol > srcrCol) { //fill right
					return FILL_RIGHT;
				}
				if (dstrCol == srcrCol && dstlCol < srclCol) { //fill left
					return FILL_LEFT;
				}
			}
		}
		return FILL_INVALID;
	}
	
	/*package*/ static Object getLibraryInstance(String key) {
		final String clsStr = Library.getProperty(key);
		if (clsStr != null) {
			try {
				final Class<?> cls = Classes.forNameByThread(clsStr);
				return cls.newInstance();
			} catch (ClassNotFoundException e) {
				//ignore
			} catch (IllegalAccessException e) {
				//ignore
			} catch (InstantiationException e) {
				//ignore
			}
		}
		return null;
	}
	
	public static Set<Ref> setRowsHidden(Sheet sheet, int tRow, int bRow, boolean hidden) {
		if (hidden) {
			for(int r = tRow; r <= bRow; ++r) {
				final Row row = BookHelper.getOrCreateRow(sheet, r);
				row.setZeroHeight(true);
			}
		} else {
			for(int r = tRow; r <= bRow; ++r) {
				final Row row = sheet.getRow(r);
				if (row != null) {
					row.setZeroHeight(false);
				}
			}
		}
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
		final Ref ref = new AreaRefImpl(tRow, 0, bRow, book.getSpreadsheetVersion().getLastColumnIndex(), refSheet);
		final Set<Ref> all = new HashSet<Ref>();
		all.add(ref);
		return all; 
	}

	public static Set<Ref> setColumnsHidden(Sheet sheet, int lCol, int rCol, boolean hidden) {
		for(int c = lCol; c <= rCol; ++c) {
			sheet.setColumnHidden(c, hidden);
		}
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = BookHelper.getRefSheet(book, sheet);
		final Ref ref = new AreaRefImpl(0, lCol, book.getSpreadsheetVersion().getLastRowIndex(), rCol, refSheet);
		final Set<Ref> all = new HashSet<Ref>();
		all.add(ref);
		return all; 
	}
	
	/*package*/ static CellRangeAddress getRepeatRowsAndColumns(Ptg[] ptgs) {
		Area3DPtg area1 = null;
		Area3DPtg area2 = null;
		for(int j = 0; j < ptgs.length; ++j) {
			Ptg ptg = ptgs[j];
			if (ptg instanceof Area3DPtg) {
				if (area1 == null) {
					area1 = (Area3DPtg) ptg;
				} else if (area2 == null) {
					area2 = (Area3DPtg) ptg;
					break;
				}
			}
		}
		if (area1 == null) {
			return new CellRangeAddress(-1, -1, -1, -1);
		} else {
			CellRangeAddress rng1 = new CellRangeAddress(area1.getFirstRow(), area1.getLastRow(), area1.getFirstColumn(), area1.getLastColumn());
			if (area2 == null) { //only area1
				if (rng1.isFullColumnRange()) { //repeat column
					return new CellRangeAddress(-1, -1, area1.getFirstColumn(), area1.getLastColumn());  
				} else if (rng1.isFullRowRange()) { //repeat row
					return new CellRangeAddress(area1.getFirstRow(), area1.getLastRow(), -1, -1);  
				}
			} else { //area1 + area2
				if (rng1.isFullColumnRange()) { //repeat column
					return new CellRangeAddress(area2.getFirstRow(), area2.getLastRow(), area1.getFirstColumn(), area1.getLastColumn());  
				} else if (rng1.isFullRowRange()) { //repeat row
					return new CellRangeAddress(area1.getFirstRow(), area1.getLastRow(), area2.getFirstColumn(), area2.getLastColumn());  
				}
			}
			return new CellRangeAddress(-1, -1, -1, -1);
		}
	}
	
	//Returns a map of (columnIndex -> Cell) of the specified row between given left column and right column
	/*package*/ static Map<Integer, Cell> copyRowCells(Row row, int lCol, int rCol) {
		if (lCol < 0 || rCol < 0) { //no cells in this row
			return null;
		}
    	final Map<Integer, Cell> cells = new HashMap<Integer, Cell>();
    	for(int c = lCol; c <= rCol; ++c) {
    		final Cell cell = row.getCell(c);
    		if (cell != null) {
    			cells.put(new Integer(c), cell);
    		}
    	}
    	return cells;
    }
	
    //20100713, henrichen@zkoss.org: copy style except border
	//create a CellStyle of the book from the given source CellStyle 
    /*package*/ static CellStyle copyFromStyleExceptBorder(Workbook book, CellStyle srcStyle) {
    	if (srcStyle.getIndex() == 0) { //default one
    		return srcStyle;
    	}
    	CellStyle dstStyle = book.createCellStyle();
		final short bb = dstStyle.getBorderBottom();
		final short tb = dstStyle.getBorderTop();
		final short lb = dstStyle.getBorderLeft();
		final short rb = dstStyle.getBorderRight();
		final short bc = dstStyle.getBottomBorderColor();
		final short tc = dstStyle.getTopBorderColor();
		final short lc = dstStyle.getLeftBorderColor();
		final short rc = dstStyle.getRightBorderColor();
		dstStyle.cloneStyleFrom(srcStyle);
		dstStyle.setBorderBottom(bb);
		dstStyle.setBorderTop(tb);
		dstStyle.setBorderLeft(lb);
		dstStyle.setBorderRight(rb);
		dstStyle.setBottomBorderColor(bc);
		dstStyle.setTopBorderColor(tc);
		dstStyle.setLeftBorderColor(lc);
		dstStyle.setRightBorderColor(rc);
		
		return dstStyle;
    }
    
    /**
     * Shifts the merged regions left or right depending on mode
     * <p>
     * @param start
     * @param end
     * @param n
     * @param horizontal
     */
    /*package*/ static List<CellRangeAddress[]> shiftMergedRegion(Sheet sheet, int tRow, int lCol, int bRow, int rCol, int n, boolean horizontal) {
        List<CellRangeAddress[]> shiftedRegions = new ArrayList<CellRangeAddress[]>();
        if (!horizontal) {
        	int start = tRow;
        	int end = bRow;
        	int dstStart = start + n;
	        //move merged regions completely if they fall within the new region boundaries when they are shifted
	        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
	        	CellRangeAddress merged = sheet.getMergedRegion(i);
	        	
	        	int firstCol = merged.getFirstColumn();
	        	int lastCol = merged.getLastColumn();
	        	if (firstCol < lCol || lastCol > rCol) { //not total cover, skip
	        		continue;
	        	}
	        	int firstRow = merged.getFirstRow();
	        	int lastRow = merged.getLastRow();
	        	
	            CellRangeAddress[] rngs = new CellRangeAddress[2]; //[0] old, [1] new
	            boolean inStart= (firstRow >= start || lastRow >= start);
	            boolean inEnd  = (firstRow <= end   || lastRow <= end);
	            
	            //not in moving area
	            if (!inStart || !inEnd) {
	            	if (n < 0 && !inStart) { //merge area in deleted area
	            		if (lastRow >= dstStart) {
	            			merged.setLastRow(dstStart - 1);
	            			if (firstRow <= merged.getLastRow()) {
	            				if (firstRow != merged.getLastRow() || lastCol != firstCol) {
	            					rngs[1] = merged;
	            				}
	            			}
	            			rngs[0] = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            				shiftedRegions.add(rngs);
	            			sheet.removeMergedRegion(i);
	            			i = i - 1; //back up now since we removed one
	            		}
	            	}
	                continue;
	            }

	            //moving area
	            if (firstRow >= start) {
	            	merged.setFirstRow(firstRow + n);
	            } else if (firstRow >= dstStart) {
	            	merged.setFirstRow(dstStart);
	            }
	            merged.setLastRow(lastRow + n);
	            if (merged.getFirstRow() != merged.getLastRow() || lastCol != firstCol) {
	            	rngs[1] = merged;
	            }
    			rngs[0] = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
	            //have to remove/add it back
				shiftedRegions.add(rngs);
	            sheet.removeMergedRegion(i);
	            i = i -1; // we have to back up now since we removed one
	        }
        } else { //20100525, henrichen@zkoss.org: add shift columns
        	int start = lCol;
        	int end = rCol;
        	int dstStart = start + n;
	        //move merged regions completely if they fall within the new region boundaries when they are shifted
	        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
	            CellRangeAddress merged = sheet.getMergedRegion(i);
	
				int firstRow = merged.getFirstRow();
				int lastRow = merged.getLastRow();
	        	if (firstRow < tRow || lastRow > bRow) { //not total cover, skip
	        		continue;
	        	}
				int firstCol = merged.getFirstColumn();
				int lastCol = merged.getLastColumn();
				
	            CellRangeAddress[] rngs = new CellRangeAddress[2]; //[0] old, [1] new
	            boolean inStart= (merged.getFirstColumn() >= start || merged.getLastColumn() >= start);
	            boolean inEnd  = (merged.getFirstColumn() <= end   || merged.getLastColumn() <= end);
	
	             //don't check if it's not within the shifted area
	            if (!inStart || !inEnd) {
	            	if (n < 0 && !inStart) { //merge area in deleted area
	            		if (lastCol >= dstStart) {
	            			merged.setLastColumn(dstStart - 1);
	            			if (firstCol <= merged.getLastColumn()) {
	            				if (firstCol != merged.getLastColumn() || lastRow != firstRow) {
	            					rngs[1] = merged;
	            				}
	            			}
	            			rngs[0] = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            				shiftedRegions.add(rngs);
	            			sheet.removeMergedRegion(i);
	            			i = i - 1; //back up now since we removed one
	            		}
	            	}
	                continue;
	            }
	
	            //moving area
	            if (firstCol >= start) {
	            	merged.setFirstColumn(firstCol + n);
	            } else if (firstCol >= dstStart) {
	            	merged.setFirstColumn(dstStart);
	            }
	            merged.setLastColumn(lastCol + n);
	            if (merged.getLastColumn() != merged.getFirstColumn() || firstRow != lastRow) {
	            	rngs[1] = merged;
	            }
    			rngs[0] = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
	            //have to remove/add it back
				shiftedRegions.add(rngs);
	            sheet.removeMergedRegion(i);
	            i = i -1; // we have to back up now since we removed one
	        }
        }
    	
        //read so it doesn't get shifted again
        Iterator<CellRangeAddress[]> iterator = shiftedRegions.iterator();
        while (iterator.hasNext()) {
            CellRangeAddress region = iterator.next()[1];
            if (region != null) {
            	sheet.addMergedRegion(region);
            }
        }
        
        return shiftedRegions;
    }

    /**
     * Shifts the merged regions left or right depending on mode
     * <p>
     * @param start
     * @param end
     * @param n
     * @param horizontal
     */
    public static List<CellRangeAddress[]> shiftBothMergedRegion(Sheet sheet, int tRow, int lCol, int bRow, int rCol, int nRow, int nCol) {
        List<CellRangeAddress[]> shiftedRegions = new ArrayList<CellRangeAddress[]>();
        //move merged regions completely if they fall within the new region boundaries when they are shifted
        final int dsttRow = tRow + nRow;
        final int dstbRow = bRow + nRow;
        final int dstlCol = lCol + nCol;
        final int dstrCol = rCol + nCol;
        
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
        	CellRangeAddress merged = sheet.getMergedRegion(i);
        	
        	int firstRow = merged.getFirstRow();
        	int lastRow = merged.getLastRow();
        	int firstCol = merged.getFirstColumn();
        	int lastCol = merged.getLastColumn();
        	if (firstCol >= lCol && lastCol <= rCol && firstRow >= tRow && lastRow <= bRow) { //source total cover
                CellRangeAddress[] rngs = new CellRangeAddress[2]; //[0] old, [1] new
       			merged.setFirstColumn(firstCol  + nCol);
       			merged.setLastColumn(lastCol + nCol);
       			merged.setFirstRow(firstRow + nRow);
       			merged.setLastRow(lastRow + nRow);
    			rngs[1] = merged;
       			rngs[0] = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
    			shiftedRegions.add(rngs);
       			sheet.removeMergedRegion(i);
       			i = i - 1; //back up now since we removed one
        		continue;
        	}
        	
        	if (firstCol >= dstlCol && lastCol <= dstrCol && firstRow >= dsttRow && lastRow <= dstbRow) { //destination total cover
                CellRangeAddress[] rngs = new CellRangeAddress[2]; //[0] old, [1] null
       			rngs[0] = merged;
    			shiftedRegions.add(rngs);
       			sheet.removeMergedRegion(i);
       			i = i - 1; //back up now since we removed one
        		continue;
        	}
        	
        	//destination partial cover (not allowed) 
        	if (firstRow <= dstbRow && lastRow >= dsttRow && firstCol <= dstrCol && lastCol >= dstlCol) {
        		throw new RuntimeException("Cannot change part of a merged cell.");
        	}
        }
    	
        //read so it doesn't get shifted again
        Iterator<CellRangeAddress[]> iterator = shiftedRegions.iterator();
        while (iterator.hasNext()) {
            CellRangeAddress region = iterator.next()[1];
            if (region != null) {
            	sheet.addMergedRegion(region);
            }
        }
        
        return shiftedRegions;
    }
    
    //20100916, henrichen@zkoss.org: assign a destination cell per the given source cell except the row and column address
	public static void assignCell(XSSFCell srcCell, XSSFCell dstCell) {
		//assign cell formats
		dstCell.setCellStyle(srcCell.getCellStyle());
		
		//assign comment
		dstCell.setCellComment(srcCell.getCellComment());
		
		//assign validation
		// TODO assign validation, refer copyValidation(srcCell, dstCell);
		
		final int cellType = srcCell.getCellType(); 
		switch(cellType) {
		case Cell.CELL_TYPE_BOOLEAN:
			dstCell.setCellValue(srcCell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			dstCell.setCellErrorValue(srcCell.getErrorCellValue());
			break;
        case Cell.CELL_TYPE_NUMERIC:
        	dstCell.setCellValue(srcCell.getNumericCellValue());
        	break;
        case Cell.CELL_TYPE_STRING:
        	dstCell.setCellValue(srcCell.getRichStringCellValue());
        	break;
        case Cell.CELL_TYPE_BLANK:
        	dstCell.setCellValue((RichTextString) null);
        	break;
        case Cell.CELL_TYPE_FORMULA:
        	dstCell.setCellFormula(srcCell.getCellFormula());
        	break;
		default:
			throw new RuntimeException("Unknown cell type:"+cellType);
		}
	}
	
	public static void setFillForegroundColor(CellStyle newCellStyle, Color xlsColor) {
		if (newCellStyle instanceof HSSFCellStyle) {
			((HSSFCellStyle)newCellStyle).setFillForegroundColor((HSSFColor)xlsColor);
		} else {
			((XSSFCellStyle)newCellStyle).setFillForegroundColor((XSSFColor)xlsColor);
		}
	}
	
	public static Color HTMLToColor(Workbook book, String color) {
		if (book instanceof HSSFWorkbook) {
			return HTMLToHSSFColor((HSSFWorkbook) book, color);
		} else {
			return HTMLToXSSFColor((XSSFWorkbook) book, color);
		}
	}
	
	private static byte[] HTMLToTriplet(String color) {
		final int offset = color.charAt(0) == '#' ? 1 : 0;
		final short red = Short.parseShort(color.substring(offset+0,offset+2), 16); //red
		final short green = Short.parseShort(color.substring(offset+2,offset+4), 16); //green
		final short blue = Short.parseShort(color.substring(offset+4, offset+6), 16); //blue
		final byte r = (byte)(red & 0xff);
		final byte g = (byte)(green & 0xff);
		final byte b = (byte)(blue & 0xff);
		return new byte[] {r, g, b};
	}
	
	private static Color HTMLToXSSFColor(XSSFWorkbook book, String color) {
		byte[] triplet = HTMLToTriplet(color);
		byte a = (byte) 0xff;
		byte r = triplet[0];
		byte g = triplet[1];
		byte b = triplet[2];
		return  new XSSFColor(new byte[] {a, r, g, b});
	}
	
	private static Color HTMLToHSSFColor(HSSFWorkbook book, String color) {
		byte[] triplet = HTMLToTriplet(color);
		byte r = triplet[0];
		byte g = triplet[1];
		byte b = triplet[2];
		short red = (short) (r & 0xff);
		short green = (short) (g & 0xff);
		short blue = (short) (b & 0xff);
		HSSFPalette palette = book.getCustomPalette();
		HSSFColor pcolor = palette != null ? palette.findColor(r, g, b) : null;
		if (pcolor != null) { //find default palette
			return pcolor;
		} else {
			final Hashtable<short[], HSSFColor> colors = HSSFColor.getTripletHash();
			HSSFColor tcolor = colors.get(new short[] {red, green, blue});
			if (tcolor != null)
				return tcolor;
			else {
				try {
					HSSFColor ncolor = palette.addColor(r, g, b);
					return ncolor;
				} catch (RuntimeException ex) {
					//return similar color if can't add new color to palette
					/**
					 * TODO: find a better solution for fix this issue
					 * 
					 * While there is no space for adding a color into palette
					 * return similar color cause return inexact color
					 */
					return palette.findSimilarColor(red, green, blue);
				}
				
			}
		}
	}
	
	public static void setLeftBorderColor(CellStyle style, Color color) {
		if (style instanceof HSSFCellStyle) {
			style.setLeftBorderColor(((HSSFColor)color).getIndex());
		} else {
			((XSSFCellStyle)style).setLeftBorderColor((XSSFColor)color);
		}
	}
	public static void setRightBorderColor(CellStyle style, Color color) {
		if (style instanceof HSSFCellStyle) {
			style.setRightBorderColor(((HSSFColor)color).getIndex());
		} else {
			((XSSFCellStyle)style).setRightBorderColor((XSSFColor)color);
		}
	}
	public static void setTopBorderColor(CellStyle style, Color color) {
		if (style instanceof HSSFCellStyle) {
			style.setTopBorderColor(((HSSFColor)color).getIndex());
		} else {
			((XSSFCellStyle)style).setTopBorderColor((XSSFColor)color);
		}
	}
	public static void setBottomBorderColor(CellStyle style, Color color) {
		if (style instanceof HSSFCellStyle) {
			style.setBottomBorderColor(((HSSFColor)color).getIndex());
		} else {
			((XSSFCellStyle)style).setBottomBorderColor((XSSFColor)color);
		}
	}
	public static Color getFontColor(Workbook book, Font font) {
		if (book instanceof XSSFWorkbook) {
			return ((XSSFFont) font).getXSSFColor();
		} else {
			return getHSSFFontColor((HSSFWorkbook) book, (HSSFFont) font);
		}
	}
	
	private static Color getHSSFFontColor(HSSFWorkbook book, HSSFFont font) {
		final short index = font.getColor() == Font.COLOR_NORMAL ? HSSFColor.AUTOMATIC.index : font.getColor();
		HSSFPalette palette = book.getCustomPalette();
		if (palette != null) {
			return palette.getColor(index);
		}
		Map<Integer, HSSFColor> indexHash = (Map<Integer, HSSFColor>) HSSFColor.getIndexHash();
		return indexHash.get(new Integer(index));
	}
	
	private static void setFontColor(Font font, Color color) {
		if (font instanceof HSSFFont) {
			((HSSFFont)font).setColor(((HSSFColor)color).getIndex());
		} else {
			((XSSFFont)font).setColor((XSSFColor)color);
		}
	}
	
	public static boolean isFreezePane(Sheet sheet) {
		return (sheet instanceof HSSFSheetImpl) ?
			((HSSFSheetImpl)sheet).isFreezePanes() : ((XSSFSheetImpl)sheet).isFreezePanes();
	}
	
	public static int getRowFreeze(Sheet sheet) {
		final PaneInformation pi = sheet.getPaneInformation();
		return pi != null ? pi.getHorizontalSplitPosition() : 0;
	}
	public static int getColumnFreeze(Sheet sheet) {
		final PaneInformation pi = sheet.getPaneInformation();
		return pi != null ? pi.getVerticalSplitPosition() : 0;
	}
	
	/*package*/ static void setDataFormat(Workbook book, CellStyle style, String formatString) {
		if (book instanceof HSSFWorkbook) {
			HSSFDataFormat df = ((HSSFWorkbook)book).createDataFormat();
			final short fmt = df.getFormat(formatString);
			style.setDataFormat(fmt);
		} else {
			XSSFDataFormat df = ((XSSFWorkbook)book).createDataFormat();
			final short fmt = df.getFormat(formatString);
			style.setDataFormat(fmt);
		}
	}
}