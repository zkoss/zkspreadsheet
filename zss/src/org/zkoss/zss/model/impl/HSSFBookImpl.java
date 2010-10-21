/* HSSFBookImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 22, 2010 7:16:56 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.awt.font.FontRenderContext;
import java.awt.font.TextAttribute;
import java.awt.font.TextLayout;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.AttributedString;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.hssf.model.InternalSheet;
import org.apache.poi.hssf.record.NameRecord;
import org.apache.poi.hssf.record.formula.Area3DPtg;
import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.record.formula.udf.AggregatingUDFFinder;
import org.apache.poi.hssf.record.formula.udf.UDFFinder;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbookHelper;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.DefaultDependencyTracker;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.formula.IStabilityClassifier;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.zkoss.lang.Classes;
import org.zkoss.lang.Library; 
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zk.ui.Component;
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
 * Implementation of {@link Book} based on HSSFWorkbook.
 * @author henrichen
 *
 */
public class HSSFBookImpl extends HSSFWorkbook implements Book {
	private final String _bookname;
	private final FormulaEvaluator _evaluator;
	private final WorkbookEvaluator _bookEvaluator;
	private final FunctionMapper _functionMapper;
	private final VariableResolver _variableResolver;
	private RefBook _refBook;
	private Books _books;
	private int _defaultCharWidth = 7; //TODO: don't know how to calculate this yet per the default font.
	private final String FUN_RESOLVER = "org.zkoss.zss.formula.FunctionResolver.class";
	private final HSSFWorkbookHelper _helper;

	public HSSFBookImpl(String bookname, InputStream is) throws IOException {
		super(is);
		_bookname = bookname;
		
		FunctionResolver resolver = (FunctionResolver) BookHelper.getLibraryInstance(FUN_RESOLVER);
		if (resolver == null) resolver = new DefaultFunctionResolver();
		_evaluator = HSSFFormulaEvaluator.create(this, NoCacheClassifier.instance, resolver.getUDFFinder()); 
		_bookEvaluator = _evaluator.getWorkbookEvaluator(); 
		_bookEvaluator.setDependencyTracker(resolver.getDependencyTracker(this));
		_functionMapper = new JoinFunctionMapper(resolver.getFunctionMapper());
		_variableResolver = new JoinVariableResolver();
		_helper = new HSSFWorkbookHelper(this);
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
					SpreadsheetVersion.EXCEL97.getLastRowIndex(), 
					SpreadsheetVersion.EXCEL97.getLastColumnIndex());
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
		return SpreadsheetVersion.EXCEL97;
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
	public int getDefaultCharWidth() {
		return _defaultCharWidth;
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

	@Override
	public CellRangeAddress getRepeatingRowsAndColumns(int sheetIndex) {
		final int nameIndex  = findExistingBuiltinNameRecordIdx(sheetIndex, NameRecord.BUILTIN_PRINT_TITLE);
		if (nameIndex == -1) {
			return new CellRangeAddress(-1, -1, -1, -1);
		}
		final NameRecord r = getNameRecord(nameIndex);
		Ptg[] ptgs = r.getNameDefinition();
		return BookHelper.getRepeatRowsAndColumns(ptgs);
	}

	//a direct copy from HSSFWorkbook#findExistingBuiltinNameRecordIdx
    private int findExistingBuiltinNameRecordIdx(int sheetIndex, byte builtinCode) {
    	final int sz = getNumberOfNames();
        for(int defNameIndex =0; defNameIndex<sz; defNameIndex++) {
            NameRecord r = _helper.getInternalWorkbook().getNameRecord(defNameIndex);
            if (r == null) {
                throw new RuntimeException("Unable to find all defined names to iterate over");
            }
            if (!r.isBuiltInName() || r.getBuiltInName() != builtinCode) {
                continue;
            }
            if (r.getSheetNumber() -1 == sheetIndex) {
                return defNameIndex;
            }
        }
        return -1;
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
    protected HSSFSheet createHSSFSheet(HSSFWorkbook workbook, InternalSheet sheet) {
    	return new HSSFSheetImpl((HSSFBookImpl)workbook, sheet);
    }
    
    @Override
    protected HSSFSheet createHSSFSheet(HSSFWorkbook workbook) {
    	return new HSSFSheetImpl((HSSFBookImpl)workbook);
    }
    
	/**
	 * Finds a font that matches the one with the supplied attributes
	 */
	public Font findFont(short boldWeight, Color color, short fontHeight, String name, boolean italic, boolean strikeout, short typeOffset, byte underline) {
		return findFont(boldWeight, ((HSSFColor)color).getIndex(), fontHeight, name, italic, strikeout, typeOffset, underline); 
	}
}
