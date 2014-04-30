/* HSSFBookImpl.java

	Purpose:
		
	Description:
		
	History:
		Mar 22, 2010 7:16:56 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashSet;
import java.util.Set;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.poi.hssf.model.InternalSheet;
import org.zkoss.poi.hssf.record.NameRecord;
import org.zkoss.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.zkoss.poi.hssf.usermodel.HSSFSheet;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbookHelper;
import org.zkoss.poi.hssf.util.HSSFColor;
import org.zkoss.poi.hssf.util.HSSFColorExt;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.PictureData;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.formula.DefaultFunctionResolver;
import org.zkoss.zss.formula.FunctionResolver;
import org.zkoss.zss.formula.NoCacheClassifier;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.BookSeries;
import org.zkoss.zss.model.NumberInputMask;
import org.zkoss.zss.model.Worksheet;

/**
 * Implementation of {@link Book} based on HSSFWorkbook.
 * @author henrichen
 *
 */
public class HSSFBookImpl extends HSSFWorkbook implements Book, BookCtrl {
	private final String _bookname;
	private final FormulaEvaluator _evaluator;
	private final WorkbookEvaluator _bookEvaluator;
	private final FunctionMapper _functionMapper;
	private final VariableResolver _variableResolver;
	private RefBook _refBook;
	private BookSeries _bookSeries;
	private int _defaultCharWidth = 7; //TODO: don't know how to calculate this yet per the default font.
	private final HSSFWorkbookHelper _helper;

	public HSSFBookImpl(String bookname, InputStream is) throws IOException {
		super(is);
		_bookname = bookname;
		
		FunctionResolver resolver = (FunctionResolver) BookHelper.getLibraryInstance(FunctionResolver.CLASS);
		if (resolver == null) resolver = new DefaultFunctionResolver();
		
		//http://tracker.zkoss.org/browse/ZSS-218
		UDFFinder udff = resolver.getUDFFinder();
		if(udff!=null){
			insertToolPack(0, udff);
		}
		_evaluator = HSSFFormulaEvaluator.create(this, NoCacheClassifier.instance,  null); // ZSS-533: not need Tolerant function anymore  
		_bookEvaluator = _evaluator.getWorkbookEvaluator(); 
		_bookEvaluator.setDependencyTracker(resolver.getDependencyTracker(this));
		_functionMapper = new JoinFunctionMapper(resolver.getFunctionMapper());
		_variableResolver = new JoinVariableResolver();
		_helper = new HSSFWorkbookHelper(this);
	}
	
	/*package*/ WorkbookEvaluator getWorkbookEvaluator() {
		return _bookEvaluator;
	}
	
	/*package*/ RefBook getOrCreateRefBook() {
		if (_refBook == null) {
			_refBook = getBookCtrl().newRefBook(this);
		}
		return _refBook;
	}
	
	/*package*/ BookSeries getBookSeries() {
		return _bookSeries;
	}
	
	/*package*/ void setBookSeries(BookSeries bookSeries) {
		_bookSeries = bookSeries;
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

	@Override
	public void deletePictureData(PictureData img) {
		super.deletePictureData(img);
	}

	//--Workbook--//
	@Override
	public void removeSheetAt(int index) {
		if (_refBook != null) {
			final String sheetname = getSheetName(index);
			_refBook.removeRefSheet(sheetname);
		}
		super.removeSheetAt(index);
		
		// ZSS-547, clear evaluator's indexes cache after deleting sheet
		// otherwise, we will get no sheet or wrong sheet
		getFormulaEvaluator().clearAllCachedResultValues();
	}

	@Override
	public void setSheetName(int index, String name) {
		final String oldsheetname = getSheetName(index);
		super.setSheetName(index, name);
		if (_refBook != null) {
			_refBook.setSheetName(oldsheetname, name);
		}
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
		if (color instanceof HSSFColorExt) {
			color = ((HSSFColorExt)color).getSimilarColor(getCustomPalette());
		}
		return findFont(boldWeight, ((HSSFColor)color).getIndex(), fontHeight, name, italic, strikeout, typeOffset, underline); 
	}
	
	private volatile BookCtrl _bookCtrl; //double-checking locking
	private BookCtrl getBookCtrl() {
		BookCtrl ctrl = _bookCtrl;
		if (ctrl == null)
			synchronized (this) {
				ctrl = _bookCtrl;
				if (ctrl == null) {
					String clsnm = Library.getProperty(BookCtrl.CLASS);
					if (clsnm != null)
						try {
							final Object o = Classes.newInstanceByThread(clsnm);
							if (!(o instanceof BookCtrl))
								throw new UiException(o.getClass().getName()+" must implement "+BookCtrl.class.getName());
							ctrl = (BookCtrl)o;
						} catch (UiException ex) {
							throw ex;
						} catch (Throwable ex) {
							throw UiException.Aide.wrap(ex, "Unable to load "+clsnm);
						}
					if (ctrl == null)
						ctrl = new BookCtrlImpl();
					_bookCtrl = ctrl;
				}
			}
		return ctrl;
	}

	@Override
	public String getShareScope() {
		return getOrCreateRefBook().getShareScope();
	}

	@Override
	public void setShareScope(String scope) {
		getOrCreateRefBook().setShareScope(scope);
	}

	@Override
	public Worksheet getWorksheetAt(int index) {
		return (Worksheet) getSheetAt(index);
	}

	@Override
	public Worksheet getWorksheet(String name) {
		return (Worksheet) getSheet(name);
	}
	
	@Override
	public boolean isDate1904() {
		return new HSSFWorkbookHelper(this).getInternalWorkbook().isUsing1904DateWindowing();
	}

	//--BookCtrl--//
	@Override
	public RefBook newRefBook(Book book) {
		return getBookCtrl().newRefBook(book);
	}
	
	@Override
	public String nextSheetId() {
		return (String) getBookCtrl().nextSheetId();
	}

	@Override
	public String nextFocusId() {
		return (String) getBookCtrl().nextFocusId();
	}

	@Override
	public void addFocus(Object focus) {
		getBookCtrl().addFocus(focus);
	}

	@Override
	public void removeFocus(Object focus) {
		getBookCtrl().removeFocus(focus);	}

	@Override
	public boolean containsFocus(Object focus) {
		return getBookCtrl().containsFocus(focus);
	}
	
	@Override
	public NumberInputMask getNumberInputMask() {
		return getBookCtrl().getNumberInputMask();
	}
}
