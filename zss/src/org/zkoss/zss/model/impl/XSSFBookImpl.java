/* XSSFBook.java

	Purpose:
		
	Description:
		
	History:
		Mar 22, 2010 7:17:28 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.awt.font.FontRenderContext;
import java.awt.font.TextAttribute;
import java.awt.font.TextLayout;
import java.io.IOException;
import java.io.InputStream;
import java.text.AttributedString;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.hssf.record.formula.udf.AggregatingUDFFinder;
import org.apache.poi.hssf.record.formula.udf.UDFFinder;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.RefBookDependencyTracker;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.impl.RefBookImpl;
import org.zkoss.zss.formula.NoCacheClassifier;
import org.zkoss.zss.formula.ZKUDFFinder;
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
	private int _defaultCharWidth;

	public XSSFBookImpl(String bookname, InputStream is) throws IOException {
		super(is);
		_bookname = bookname;
		_evaluator = XSSFFormulaEvaluator.create(this, new NoCacheClassifier(), 
				new AggregatingUDFFinder(UDFFinder.DEFAULT, ZKUDFFinder.instance));
		_bookEvaluator = _evaluator.getWorkbookEvaluator(); 
		_bookEvaluator.setDependencyTracker(new RefBookDependencyTracker(this));
		_functionMapper =  new JoinFunctionMapper();
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

    private void copyAttributes(Font font, AttributedString str, int startIdx, int endIdx) {
        str.addAttribute(TextAttribute.FAMILY, font.getFontName(), startIdx, endIdx);
        str.addAttribute(TextAttribute.SIZE, new Float(font.getFontHeightInPoints()));
        if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD) str.addAttribute(TextAttribute.WEIGHT, TextAttribute.WEIGHT_BOLD, startIdx, endIdx);
        if (font.getItalic() ) str.addAttribute(TextAttribute.POSTURE, TextAttribute.POSTURE_OBLIQUE, startIdx, endIdx);
        if (font.getUnderline() == Font.U_SINGLE ) str.addAttribute(TextAttribute.UNDERLINE, TextAttribute.UNDERLINE_ON, startIdx, endIdx);
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
}
