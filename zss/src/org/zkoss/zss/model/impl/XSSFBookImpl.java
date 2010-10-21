/* XSSFBook.java

	Purpose:
		
	Description:
		
	History:
		Mar 22, 2010 7:17:28 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.hssf.record.formula.Area3DPtg;
import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.lang.Classes;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.impl.RefBookImpl;
import org.zkoss.zss.formula.DefaultFunctionResolver;
import org.zkoss.zss.formula.FunctionResolver;
import org.zkoss.zss.formula.NoCacheClassifier;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Books;

/**
 * Implementation of {@link Book} based on XSSFWorkbook.
 * @author henrichen
 *
 */
public class XSSFBookImpl extends XSSFWorkbook implements Book {
	private final String _bookname;
	private final FormulaEvaluator _evaluator;
	private final WorkbookEvaluator _bookEvaluator;
	private final FunctionMapper _functionMapper;
	private final VariableResolver _variableResolver;
	private RefBook _refBook;
	private Books _books;
	private int _defaultCharWidth = 7; //TODO: don't know how to calculate this yet per the default font.
	private final String FUN_RESOLVER = "org.zkoss.zss.formula.FunctionResolver.class";
	
	//override the XSSFSheet Relation
	static {
		Field fd = null;
		try {
			fd = Classes.getAnyField(XSSFRelation.class, "_cls");
		} catch (NoSuchFieldException e) {
			throw new RuntimeException(e);
		}
		final boolean old = fd.isAccessible();
		try {
			fd.setAccessible(true);
			fd.set(XSSFRelation.WORKSHEET, XSSFSheetImpl.class); //Use the new XSSFSheet implementation 
		} catch (IllegalAccessException e) {
			throw new RuntimeException(e);
		} finally {
			fd.setAccessible(old);
		}
	}

	public XSSFBookImpl(String bookname, InputStream is) throws IOException {
		super(is);
		_bookname = bookname;
		FunctionResolver resolver = (FunctionResolver) BookHelper.getLibraryInstance(FUN_RESOLVER);
		if (resolver == null) resolver = new DefaultFunctionResolver();
		_evaluator = XSSFFormulaEvaluator.create(this, NoCacheClassifier.instance, resolver.getUDFFinder()); 
		_bookEvaluator = _evaluator.getWorkbookEvaluator(); 
		_bookEvaluator.setDependencyTracker(resolver.getDependencyTracker(this));
		_functionMapper = new JoinFunctionMapper(resolver.getFunctionMapper());
		_variableResolver = new JoinVariableResolver();
	}
	
	/*package*/ WorkbookEvaluator getWorkbookEvaluator() {
		return _bookEvaluator;
	}
	
	/*package*/ RefBook getRefBook() {
		return _refBook;
	}
	
	/*package*/ RefBook getOrCreateRefBook() {
		if (_refBook == null) {
			_refBook = new RefBookImpl(_bookname,
					SpreadsheetVersion.EXCEL2007.getLastRowIndex(), 
					SpreadsheetVersion.EXCEL2007.getLastColumnIndex());
		}
		return _refBook;
	}
	
	/*package*/ Books getBooks() {
		return _books;
	}
	
	/*package*/ void setBooks(Books books) {
		_books = books;
	}
	
	/*package*/ VariableResolver getVariableResolver() {
		return _variableResolver;
	}
	
	/*package*/ FunctionMapper getFunctionMapper() {
		return _functionMapper;
	}
	
	//--Book--//
	@Override
	public SpreadsheetVersion getSpreadsheetVersion() {
		return SpreadsheetVersion.EXCEL2007;
	}
	
	@Override
	public String getBookName() {
		return _bookname;
	}
	
	@Override
	public FormulaEvaluator getFormulaEvaluator() {
		return _evaluator;
	}

	@Override
	public void addFunctionMapper(FunctionMapper mapper) {
		((JoinFunctionMapper)getFunctionMapper()).addFunctionMapper(mapper);
	}

	@Override
	public void addVariableResolver(VariableResolver resolver) {
		((JoinVariableResolver)getVariableResolver()).addVariableResolver(resolver);
	}

	@Override
	public void removeFunctionMapper(FunctionMapper mapper) {
		((JoinFunctionMapper)getFunctionMapper()).removeFunctionMapper(mapper);
	}

	@Override
	public void removeVariableResolver(VariableResolver resolver) {
		((JoinVariableResolver)getVariableResolver()).removeVariableResolver(resolver);
	}

	@Override
	public void subscribe(EventListener listener) {
		getOrCreateRefBook().subscribe(listener);
	}

	@Override
	public void unsubscribe(EventListener listener) {
		getOrCreateRefBook().unsubscribe(listener);
	}
	
	@Override
	public Font getDefaultFont() {
		return getFontAt((short)0);
	}
	
	@Override
	public void setDefaultFont(Font font) {
		final Font defFont = getDefaultFont();
		defFont.setBoldweight(font.getBoldweight());
		defFont.setCharSet(font.getCharSet());
		defFont.setColor(font.getColor());
		defFont.setFontHeight(font.getFontHeight());
		defFont.setFontName(font.getFontName());
		defFont.setItalic(font.getItalic());
		defFont.setStrikeout(font.getStrikeout());
		defFont.setTypeOffset(font.getTypeOffset());
		defFont.setUnderline(font.getUnderline());
		
		//TODO: recalic _defaultCharWidth
	}
	
	@Override
	public int getDefaultCharWidth() {
		return _defaultCharWidth;
	}
	
	@Override 
	public void notifyChange(String[] variables) {
		final RefBook refBook = getOrCreateRefBook();
		final Set<Ref> all = new HashSet<Ref>();
		final Set<Ref> last = new HashSet<Ref>();
		for(String name : variables) {
			final Set<Ref>[] refs = refBook.getBothDependents(name);
			if (refs != null) {
				last.addAll(refs[0]);
				all.addAll(refs[1]);
			}
		}
		BookHelper.reevaluateAndNotify(this, last, all);
	}
	
	//--Workbook--//
	@Override
	public void removeSheetAt(int index) {
		final String sheetname = getSheetName(index);
		_refBook.removeRefSheet(sheetname);
		super.removeSheetAt(index);
	}

	@Override
	public void setSheetName(int index, String name) {
		final String oldsheetname = getSheetName(index);
		_refBook.setSheetName(oldsheetname, name);
		super.setSheetName(index, name);
	}

	@Override
	public CellRangeAddress getRepeatingRowsAndColumns(int sheetNumber) {
		final XSSFName name = getBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetNumber);
		if (name == null) {
			return new CellRangeAddress(-1, -1, -1, -1);
		} else {
			final String formula = name.getRefersToFormula();
			final Ptg[] ptgs = FormulaParser.parse(formula, XSSFEvaluationWorkbook.create(this), FormulaType.NAMEDRANGE, name.getSheetIndex());
			return BookHelper.getRepeatRowsAndColumns(ptgs);
		}
	}
    private XSSFName getBuiltInName(String builtInCode, int sheetNumber) {
    	final int sz  = getNumberOfNames();
    	for (int j = 0; j < sz; ++j) {
    		XSSFName name = getNameAt(j);
            if (name.getNameName().equalsIgnoreCase(builtInCode) && name.getSheetIndex() == sheetNumber) {
                return name;
            }
        }
        return null;
    }

	/**
	 * Finds a font that matches the one with the supplied attributes
	 */
	public XSSFFont findFont(short boldWeight, Color color, short fontHeight, String name, boolean italic, boolean strikeout, short typeOffset, byte underline) {
		for (XSSFFont font : getStylesSource().getFonts()) {
			final XSSFColor fontColor = font.getXSSFColor();
			if (	(font.getBoldweight() == boldWeight)
					&& (color == fontColor || color != null && color.equals(fontColor))
					&& font.getFontHeight() == fontHeight
					&& font.getFontName().equals(name)
					&& font.getItalic() == italic
					&& font.getStrikeout() == strikeout
					&& font.getTypeOffset() == typeOffset
					&& font.getUnderline() == underline)
			{
				return font;
			}
		}
		return null;
	}
}
