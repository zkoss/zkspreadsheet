/* BookHelper.java

	Purpose:
		
	Description:
		
	History:
		Mar 24, 2010 5:42:58 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.Hashtable;
import java.util.List;
import java.util.Locale;
import java.util.Set;

import org.apache.poi.hssf.record.CellValueRecordInterface;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.aggregates.FormulaRecordAggregate;
import org.apache.poi.hssf.record.formula.AreaPtgBase;
import org.apache.poi.hssf.record.formula.ArrayPtg;
import org.apache.poi.hssf.record.formula.AttrPtg;
import org.apache.poi.hssf.record.formula.ExpPtg;
import org.apache.poi.hssf.record.formula.FormulaShifter;
import org.apache.poi.hssf.record.formula.OperandPtg;
import org.apache.poi.hssf.record.formula.OperationPtg;
import org.apache.poi.hssf.record.formula.ParenthesisPtg;
import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.record.formula.RefPtgBase;
import org.apache.poi.hssf.record.formula.ScalarConstantPtg;
import org.apache.poi.hssf.record.formula.TblPtg;
import org.apache.poi.hssf.record.formula.UnknownPtg;
import org.apache.poi.hssf.usermodel.DummyHSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaParsingWorkbook;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.ErrorConstants;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
import org.zkoss.zss.engine.impl.DependencyTrackerHelper;
import org.zkoss.zss.engine.impl.MergeChange;
import org.zkoss.zss.engine.impl.RefSheetImpl;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Books;

/**
 * Internal User only. Helper class to handle the two version of Books.
 * @author henrichen
 */
public final class BookHelper {
	public static final String AUTO_COLOR = "AUTO_COLOR";
	
	public static RefBook getRefBook(Book book) {
		return book instanceof HSSFBookImpl ? 
			((HSSFBookImpl)book).getRefBook():
			((XSSFBookImpl)book).getRefBook();
	}
	
	public static RefBook getOrCreateRefBook(Book book) {
		return book instanceof HSSFBookImpl ? 
			((HSSFBookImpl)book).getOrCreateRefBook():
			((XSSFBookImpl)book).getOrCreateRefBook();
	}
	
	public static RefSheet getRefSheet(Book book, Sheet sheet) {
		final RefBook refBook = getRefBook(book);
		return refBook.getRefSheet(sheet.getSheetName());
	}
	
	public static Books getBooks(Book book) {
		return book instanceof HSSFBookImpl ? 
				((HSSFBookImpl)book).getBooks():
				((XSSFBookImpl)book).getBooks();
	}
	
	public static void setBooks(Book book, Books books) {
		if (book instanceof HSSFBookImpl) 
			((HSSFBookImpl)book).setBooks(books);
		else
			((XSSFBookImpl)book).setBooks(books);
	}
	
	public static Book getBook(Book book, String bookname) {
		final Books books = BookHelper.getBooks(book);
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
		for(Ref ref : all) {
			final RefSheet refSheet = ref.getOwnerSheet();
			final Book bookTarget = BookHelper.getBook(book, refSheet);
			final Cell cell = getCell(book, ref.getTopRow(), ref.getLeftCol(), refSheet);
			if (cell != null)
				bookTarget.getFormulaEvaluator().notifySetFormula(cell); 
		}
	}
	
	/*package*/ static void reevaluate(Book book, Set<Ref> last) {
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
	
	/*package*/ static void notifyCellChanges(Book book, Set<Ref> all) {
		for(Ref ref : all) {
			final RefSheet refSheet = ref.getOwnerSheet();
			final RefBook refBook = refSheet.getOwnerBook();
			refBook.publish(new SSDataEvent(SSDataEvent.ON_CONTENTS_CHANGE, ref, SSDataEvent.MOVE_NO));
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
	
	public static String formatRichText(Book book, RichTextString rstr, List<int[]> indexes) {
		final String raw = rstr.getString();
		final int runs = rstr.numFormattingRuns();
		if (runs > 0) {
			final int len = rstr.length();
			StringBuffer sb = new StringBuffer(len + runs*96);
			int e = 0;
			for(int b = 0, j = 0; j < runs; ++j) {
				if (j > 0) {
					sb.append("</span>");
				}
				final int sbb = sb.length();
				e = rstr.getIndexOfFormattingRun(j);
				sb.append(raw.substring(b, e));
				if (indexes != null) {
					final int sbe = sb.length();
					indexes.add(new int[] {sbb, sbe});
				}
				b = e;
				final Font font = getFont(book, rstr, j);
				sb.append("<span style=\"")
					.append(getFontCSSStyle(book, font))
					.append("\">");
			}
			if (e < len) {
				final int sbb = sb.length();
				sb.append(raw.substring(e));
				if (indexes != null) {
					final int sbe = sb.length();
					indexes.add(new int[] {sbb, sbe});
				}
				sb.append("</span>");
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
		
		String textColor = BookHelper.getFontColor(book, font);
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
	
	public static String getFontColor(Book book, Font font) {
		if (font instanceof XSSFFont) {
			final XSSFFont f = (XSSFFont) font;
			final XSSFColor color = f.getXSSFColor();
			final byte[] triplet = color.getRgb();
			return triplet == null ? null : color.isAuto() ? AUTO_COLOR : 
				"#"+ toHex(triplet[1])+ toHex(triplet[2])+ toHex(triplet[3]);
		} else {
			return indexToRGB(book, font.getColor());
		}
	}
	
	@SuppressWarnings("unchecked")
	public static String indexToRGB(Book book, int index) {
		HSSFPalette palette = ((HSSFWorkbook)book).getCustomPalette();
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
		byte red = Byte.parseByte(color.substring(1,3), 16); //red
		byte green = Byte.parseByte(color.substring(3,5), 16); //green
		byte blue = Byte.parseByte(color.substring(5), 16); //blue
		HSSFColor pcolor = palette.findColor(red, green, blue);
		if (pcolor != null) { //find default palette
			return pcolor.getIndex();
		} else {
			final Hashtable<short[], HSSFColor> colors = HSSFColor.getTripletHash();
			HSSFColor tcolor = colors.get(new short[] {red, green, blue});
			if (tcolor != null)
				return tcolor.getIndex();
			else {
				HSSFColor ncolor = palette.addColor(red, green, blue);
				return ncolor.getIndex();
			}
		}
	}
	
	private static String toHex(int num) {
		final String hex = Integer.toHexString(num);
		return hex.length() == 1 ? "0"+hex : hex; 
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

/*		String result = null;
		switch(cellType) {
		case Cell.CELL_TYPE_BLANK:
			result = "";
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			result = cell.getBooleanCellValue() ? "TRUE" : "FALSE";
			break;
		case Cell.CELL_TYPE_ERROR:
			result = ErrorConstants.getText(cell.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			final Book book = (Book)cell.getSheet().getWorkbook();
			final CellValue cv = BookHelper.evaluate(book, cell);
			result = getTextByCellValue(cv);
			break;
		case Cell.CELL_TYPE_NUMERIC:
			result = NumberToTextConverter.toText(cell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING:
			return cell.getRichStringCellValue();
		default:
			throw new UiException("Unknown cell type:"+cellType);
		}
*/		
		return cell instanceof HSSFCell ? new HSSFRichTextString(result) : new XSSFRichTextString(result);
	}
	
	private static Object getValueByCellValue(CellValue cellValue) {
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
	
	private static String getTextByCellValue(CellValue cellValue) {
		final int cellType = cellValue.getCellType();
		switch(cellType) {
		case Cell.CELL_TYPE_BLANK:
			return "";
		case Cell.CELL_TYPE_BOOLEAN:
			return cellValue.getBooleanValue() ? "TRUE" : "FALSE";
		case Cell.CELL_TYPE_ERROR:
			return ErrorConstants.getText(cellValue.getErrorValue());
		case Cell.CELL_TYPE_NUMERIC:
			return NumberToTextConverter.toText(cellValue.getNumberValue());
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
			return rtstr == null ? "" : rtstr.getString(); 
		default:
			throw new UiException("Unknown cell type:"+cellType);
		}
	}
	
	public static Set<Ref>[] removeCell(Sheet sheet, int rowIndex, int colIndex) {
		final Cell cell = getCell(sheet, rowIndex, colIndex);
		if (cell != null) {
			//remove formula cell and create a blank one
			removeFormula(cell);
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
		removeFormula(cell);
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
		removeFormula(cell);
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
		removeFormula(cell);
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
		removeFormula(cell);
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
		removeFormula(cell);
		//set value into cell model
		cell.setCellFormula(value);
		//notify to update cache
		return getBothDependents(cell); 
	}
	
	public static Set<Ref>[] setCellValue(Cell cell, String value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_STRING, value))
			return null;
		//remove formula cell and crete a blank one
		removeFormula(cell);
		//set value into cell model
		cell.setCellValue(value);
		//notify to update cache
		return getBothDependents(cell); 
	}
	
	public static Set<Ref>[] setCellValue(Cell cell, RichTextString value) {
		//same type and value, return!
		if (sameTypeAndValue(cell, Cell.CELL_TYPE_STRING, value))
			return null;
		//remove formula cell and crete a blank one
		removeFormula(cell);
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
	
	private static Set<Ref>[] getBothDependents(Cell cell) {
		final Sheet sheet = cell.getSheet();
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		
		//clear/update formula cache
		book.getFormulaEvaluator().notifySetFormula(cell); 
		//get affected dependents(last, all)
		return ((RefSheetImpl)refSheet).getBothDependents(cell.getRowIndex(), cell.getColumnIndex());
	}
	
	//formula cell -> non-formula cell
	private static void removeFormula(Cell cell) {
		//Was formula, shall maintain the formula cache and dependency 
		if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			//clear the formula reference dependency
			final int rowIndex = cell.getRowIndex(); 
			final int colIndex = cell.getColumnIndex();
			final Sheet sheet = cell.getSheet();
			final Book book = (Book) sheet.getWorkbook();
			final RefSheet refSheet = getRefSheet(book, sheet);
			final Ref ref = refSheet.getRef(rowIndex, colIndex, rowIndex, colIndex);
			ref.removeAllPrecedents();
			
			//remove formula from the cell
			cell.setCellFormula(null);
			
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
					final Date val = parseToDate(txt);
					return val != null ? 
						new Object[] {new Integer(Cell.CELL_TYPE_NUMERIC), val}: //date
						new Object[] {new Integer(Cell.CELL_TYPE_STRING), txt}; //string
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

	private static String[] DATE_TIME_PATTERN  = new String[] {
		"M/d/yy", "M-d-yy", "M/d", "M-d", "yyyy/M/d", "yyyy-M-d", "d-MMM-yy",
		"MMMM-yy", "MMM-yy", "MMMM d, yyyy", "M/d/yy h:mm a", "M/d/yy H:mm", "MMMM",
		"h:m:s a", "H:m:s", "h:m a", "H:m",
	};
	
	//TODO, support locale other than US
	private static Date parseToDate(String txt) {
		for(int j = 0; j < DATE_TIME_PATTERN.length; ++j) {
			try {
				final DateFormat df = new SimpleDateFormat(DATE_TIME_PATTERN[j], Locale.US);
				return df.parse(txt);
			} catch (ParseException ex) {
				//ignore, try next pattern
			}
		}
		return null;
	}
	
	public static List<Picture> getPictures(Sheet sheet) {
		if (sheet instanceof HSSFSheet) {
			final HSSFSheet sheet0 = (HSSFSheet) sheet;
			final HSSFPatriarch patriarch = sheet0.getDrawingPatriarch();
			if (patriarch != null) {
				final List<HSSFShape> shapes = patriarch.getChildren();
				final List<Picture> pictures = new ArrayList<Picture>(shapes.size());
				for(HSSFShape shape : shapes) {
					if (shape instanceof HSSFPicture) {
						pictures.add((Picture) shape);
					}
				}
				return pictures;
			}
			return null;
		} else if (sheet instanceof XSSFSheet) {
			//TODO support XSSFSheet picture (see XSSFDrawing)
			final XSSFSheet sheet0 = (XSSFSheet) sheet;
			final XSSFDrawing drawing = sheet0.createDrawingPatriarch();
			throw new UnsupportedOperationException("Picture in XSSFSheet not supported yet: "+sheet);
		}
		throw new UnsupportedOperationException("Unknow sheet type(must be either HSSFSheet or XSSFSheet: "+sheet);
	}
	
	public static String getPictureKey(Picture picture) {
		if (picture instanceof HSSFPicture) {
			return ""+((HSSFPicture)picture).getPictureIndex();
		} else if (picture instanceof XSSFPicture) {
			//TODO support XSSFPicture.getData()
			throw new UnsupportedOperationException("XSSFPicture is not supported yet!");
		}
		throw new UnsupportedOperationException("Unknow sheet type(must be either HSSFPicture or XSSFPicture: "+picture);
	}
	
	public static PictureData getData(List<? extends PictureData> datas, Picture picture) {
		if (datas != null){
			if (picture instanceof HSSFPicture) {
				final int index = ((HSSFPicture)picture).getPictureIndex();
				if (datas.size() > index) {
					final PictureData data = datas.get(index);
					if (data != null)
						return data;
				}
			} else if (picture instanceof XSSFPicture) {
				//TODO support XSSFPicture.getData()
				throw new UnsupportedOperationException("XSSFPicture is not supported yet!");
			}
		}
		return null;
	}
	
	public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		return row;
	}
	
	public static Set<Ref>[] copyCell(Cell cell, Sheet sheet, int rowIndex, int colIndex) {
		//TODO now assume same workbook in copying CellStyle. if different workbook (xls to xls, use CellStyle.cloneStyleFrom())
		final Cell dstCell = getOrCreateCell(sheet, rowIndex, colIndex);
		dstCell.setCellStyle(cell.getCellStyle());
		final int cellType = cell.getCellType(); 
		switch(cellType) {
		case Cell.CELL_TYPE_BOOLEAN:
			return setCellValue(dstCell, cell.getBooleanCellValue());
		case Cell.CELL_TYPE_ERROR:
			return setCellValue(dstCell, cell.getErrorCellValue());
        case Cell.CELL_TYPE_NUMERIC:
        	return setCellValue(dstCell, cell.getNumericCellValue());
        case Cell.CELL_TYPE_STRING:
        	return setCellValue(dstCell, cell.getRichStringCellValue());
        case Cell.CELL_TYPE_BLANK:
        	return setCellValue(dstCell, (String) null);
        case Cell.CELL_TYPE_FORMULA:
        	return copyCellFormula(dstCell, cell);
		default:
			throw new UiException("Unknown cell type:"+cellType);
		}
	}
	
	private static Set<Ref>[] copyCellFormula(Cell dstCell, Cell srcCell) {
		//same type and value, return!
//		if (sameTypeAndValue(dstCell, Cell.CELL_TYPE_FORMULA, value))
//			return null;
		
		//remove formula cell and create a blank one
		removeFormula(dstCell);
		
		//set value into cell model
		final Ptg[] dstPtgs = offsetPtgs(dstCell, srcCell);
		setCellPtgs(dstCell, dstPtgs);
		evaluate((Book)dstCell.getSheet().getWorkbook(), dstCell);
		
		//notify to update cache
		return getBothDependents(dstCell); 
	}
	
	private static void setCellPtgs(Cell cell, Ptg[] ptgs) {
        if (cell instanceof HSSFCell) {
        	DummyHSSFCell dummyCell = new DummyHSSFCell((HSSFCell)cell);
        	//tricky! must be after the dummyCell construction or the under aggregate record will not 
        	//be consistent in sheet and cell
            cell.setCellType(Cell.CELL_TYPE_FORMULA); 
        	FormulaRecordAggregate agg = (FormulaRecordAggregate) dummyCell.getRecord();
        	FormulaRecord frec = agg.getFormulaRecord();
        	frec.setOptions((short) 2);
        	frec.setValue(0);

        	//only set to default if there is no extended format index already set
        	if (agg.getXFIndex() == (short)0) {
        		agg.setXFIndex((short) 0x0f);
        	}
        	agg.setParsedExpression(ptgs);
        	return;
        }
		throw new UiException("Unknown cell class:"+cell);
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
		CellValueRecordInterface vr = new DummyHSSFCell(cell).getRecord();
		if (!(vr instanceof FormulaRecordAggregate)) {
			throw new IllegalArgumentException("Not a formula cell");
		}
		FormulaRecordAggregate fra = (FormulaRecordAggregate) vr;
		return fra.getFormulaRecord().getParsedExpression();
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
		
		for(int j = 0; j < ptglen; ++j) {
			final Ptg srcPtg = srcPtgs[j];
			final Ptg dstPtg = offsetPtg(srcSheet, srcPtg, dstSheet, startRow, offRow, startCol, offCol);
			dstPtgs[j] = dstPtg;
		}
		return dstPtgs;
	}
	
	private static Ptg offsetPtg(Sheet srcSheet, Ptg srcPtg, Sheet dstSheet, int startRow, int offRow, int startCol, int offCol) {
		//TODO shall handle different sheet, different book case
		
		//ScalarConstantPtg
		if (srcPtg instanceof ScalarConstantPtg) { //constant
			return srcPtg;
		}
		
		//OperandPtg
		if (srcPtg instanceof OperandPtg) {
			//TODO if column or row over the maximum limit, shall return #Ref! error ptg
			final Ptg dstPtg = ((OperandPtg)srcPtg).copy(); 
			if (srcPtg instanceof RefPtgBase) {
				final RefPtgBase srcb = (RefPtgBase) srcPtg;
				final RefPtgBase dstb = (RefPtgBase) dstPtg;
				if (offRow > 0) {
					if (startRow <= srcb.getRow()) {
						dstb.setRow(srcb.getRow()+offRow);
					}
				} else if (offRow < 0) {
					throw new UiException("offRow < 0 not implemented yet");
				}
				if (offCol > 0) {
					if (startCol <= srcb.getColumn()) {
						dstb.setColumn(srcb.getColumn()+offCol);
					}
				} else if (offCol < 0) {
					throw new UiException("offCol < 0 not implemented yet");
				}
				dstb.setRowRelative(srcb.isRowRelative());
				dstb.setColRelative(srcb.isColRelative());
			} else if (srcPtg instanceof AreaPtgBase) {
				final AreaPtgBase srcb = (AreaPtgBase) srcPtg;
				final AreaPtgBase dstb = (AreaPtgBase) dstPtg;
				if (offRow > 0) {
					if (startRow <= srcb.getFirstRow()) {
						dstb.setFirstRow(srcb.getFirstRow()+offRow);
					}
				} else if (offRow < 0) {
					throw new UiException("offRow < 0 not implemented yet");
				}
				if (offRow > 0) {
					if (startRow <= srcb.getLastRow()) {
						dstb.setLastRow(srcb.getLastRow()+offRow);
					}
				} else if (offRow < 0) {
					throw new UiException("offRow < 0 not implemented yet");
				}
				if (offCol > 0) {
					if (startCol <= srcb.getFirstColumn()) {
						dstb.setFirstColumn(srcb.getFirstColumn()+offCol);
					}
				} else if (offCol < 0) {
					throw new UiException("offCol < 0 not implemented yet");
				}
				if (offCol > 0) {
					if (startCol <= srcb.getLastColumn()) {
						dstb.setLastColumn(srcb.getLastColumn()+offCol);
					}
				} else if (offCol < 0) {
					throw new UiException("offCol < 0 not implemented yet");
				}
				dstb.setFirstRowRelative(srcb.isFirstRowRelative());
				dstb.setLastRowRelative(srcb.isFirstRowRelative());
				dstb.setFirstColRelative(srcb.isFirstColRelative());
				dstb.setLastColRelative(srcb.isFirstColRelative());
			}
			return dstPtg;
		}
		
		//ControlPtg
//		if (srcPtg instanceof ControlPtg) {
			if (srcPtg instanceof ParenthesisPtg) {
				return srcPtg;
			}
			
			if (srcPtg instanceof AttrPtg) { //TODO not sure if we can return srcPtg
				return srcPtg;
			}
			
			if (srcPtg instanceof ExpPtg) { //TODO row, col inside, not sure if we can return srcPtg
				return srcPtg;
			}
			
			if (srcPtg instanceof TblPtg) { //TODO row, col inside, not sure if we can return srcPtg
				return srcPtg;
			}
//		}
		
		//OperationPtg
		if (srcPtg instanceof OperationPtg) { 
			return srcPtg;
		}
		
		//ArrayPtg
		if (srcPtg instanceof ArrayPtg) { 
			//TODO not sure what will happen copy Array to another cell
			return srcPtg;
		}
		
		//UnknownPtg
		if (srcPtg instanceof UnknownPtg) {
			return srcPtg;
		}
		
		throw new UiException("Unknown Ptg :"+srcPtg);
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
	
	public static ChangeInfo insertRows(Sheet sheet, int startRow, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertRows(startRow, num);
		final int lastRowNum = sheet.getLastRowNum();
		if (startRow > lastRowNum) {
			return null;
		}
		final List<CellRangeAddress> shiftedRanges = ((HSSFSheet)sheet).shiftRowsOnly(startRow, lastRowNum, num, true, false, true, false);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges, num, true);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		shiftFormulas(all, sheet, startRow, refSheet.getOwnerBook().getMaxrow(), num, true);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	public static ChangeInfo deleteRows(Sheet sheet, int startRow, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteRows(startRow, num);
		final int lastRowNum = sheet.getLastRowNum();
		final int startRow0 = startRow + num;
		if (startRow > lastRowNum) {
			return null;
		}
		final List<CellRangeAddress> shiftedRanges = ((HSSFSheet)sheet).shiftRowsOnly(startRow0, lastRowNum, -num, true, false, true, true);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges, -num, true);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		shiftFormulas(all, sheet, startRow0, refSheet.getOwnerBook().getMaxrow(), -num, true);
		
		return new ChangeInfo(last, all, changeMerges);
	}

	public static ChangeInfo insertColumns(Sheet sheet, int startCol, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.insertColumns(startCol, num);
		final List<CellRangeAddress> shiftedRanges = ((HSSFSheet)sheet).shiftColumnsOnly(startCol, -1, num, true, false, true, false);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges, num, false);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		shiftFormulas(all, sheet, startCol, refSheet.getOwnerBook().getMaxcol(), num, false);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	public static ChangeInfo deleteColumns(Sheet sheet, int startCol, int num) {
		final Book book = (Book) sheet.getWorkbook();
		final RefSheet refSheet = getRefSheet(book, sheet);
		final Set<Ref>[] refs = refSheet.deleteColumns(startCol, num);
		final int startCol0 = startCol + num;
		final List<CellRangeAddress> shiftedRanges = ((HSSFSheet)sheet).shiftColumnsOnly(startCol0, -1, -num, true, false, true, true);
		final List<MergeChange> changeMerges = prepareChangeMerges(refSheet, shiftedRanges, -num, false);
		final Set<Ref> last = refs[0];
		final Set<Ref> all = refs[1];
		shiftFormulas(all, sheet, startCol0, refSheet.getOwnerBook().getMaxcol(), -num, false);
		
		return new ChangeInfo(last, all, changeMerges);
	}
	
	private static List<MergeChange> prepareChangeMerges(RefSheet sheet, List<CellRangeAddress> shiftedRanges, int num, boolean isRow) {
		final List<MergeChange> changeMerges = new ArrayList<MergeChange>();
		for(CellRangeAddress rng : shiftedRanges) {
			int tRow = rng.getFirstRow();
			int lCol = rng.getFirstColumn();
			int bRow = rng.getLastRow();
			int rCol = rng.getLastColumn();
			final Ref ref = new AreaRefImpl(tRow, lCol, bRow, rCol, sheet);
			if (isRow) {
				tRow -= num;
				bRow -= num;
			} else {
				lCol -= num;
				rCol -= num;
			}
			final Ref orgRef = new AreaRefImpl(tRow, lCol, bRow, rCol, sheet);
			changeMerges.add(new MergeChange(orgRef, ref));
		}
		return changeMerges;
	}
	
	private static void shiftFormulas(Set<Ref> all, Sheet sheet, int startRow, int endRow, int n, boolean isRow) {
		final int moveSheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
        final FormulaShifter shifter = isRow ? 
        		FormulaShifter.createForRowShift(moveSheetIndex, startRow, endRow, n):
				FormulaShifter.createForColumnShift(moveSheetIndex, startRow, endRow, n);
		for (Ref ref : all) {
			final int tRow = ref.getTopRow();
			final int lCol = ref.getLeftCol();
			final Sheet srcSheet = getSheet(sheet, ref.getOwnerSheet());
			final Book srcBook = (Book) srcSheet.getWorkbook();
			final Cell srcCell = getCell(srcSheet, tRow, lCol);
			if (srcCell != null && srcCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
				final int sheetIndex = srcBook.getSheetIndex(srcSheet);
				if (srcCell instanceof HSSFCell) {
					HSSFShiftFormulas(shifter, sheetIndex, srcCell);
				} else {
					throw new UiException("ShiftFormula of Excel 2007(.xlsx) file in not implemented yet");
				}
			}
		}
	}
	
	private static void HSSFShiftFormulas(FormulaShifter shifter, int sheetIndex, Cell cell) {
		CellValueRecordInterface vr = new DummyHSSFCell((HSSFCell)cell).getRecord();
		if (!(vr instanceof FormulaRecordAggregate)) {
			throw new IllegalArgumentException("Not a formula cell");
		}
		FormulaRecordAggregate fra = (FormulaRecordAggregate) vr;
		Ptg[] ptgs = fra.getFormulaRecord().getParsedExpression();
		if (shifter.adjustFormula(ptgs, sheetIndex)) {
			fra.setParsedExpression(ptgs);
		}
	}
}