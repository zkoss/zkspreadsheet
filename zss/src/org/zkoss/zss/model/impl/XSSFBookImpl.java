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
import java.util.List;
import java.util.Set;

import org.openxmlformats.schemas.officeDocument.x2006.docPropsVTypes.CTVariant;
import org.openxmlformats.schemas.officeDocument.x2006.docPropsVTypes.CTVector;
import org.openxmlformats.schemas.officeDocument.x2006.extendedProperties.CTProperties;
import org.openxmlformats.schemas.officeDocument.x2006.extendedProperties.CTVectorLpstr;
import org.openxmlformats.schemas.officeDocument.x2006.extendedProperties.CTVectorVariant;
import org.zkoss.lang.Classes;
import org.zkoss.lang.Library;
import org.zkoss.poi.POIXMLProperties;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.SheetNameFormatter;
import org.zkoss.poi.ss.formula.WorkbookEvaluator;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.PictureData;
import org.zkoss.poi.ss.usermodel.PivotCache;
import org.zkoss.poi.ss.util.AreaReference;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFColor;
import org.zkoss.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.zkoss.poi.xssf.usermodel.XSSFFont;
import org.zkoss.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.zkoss.poi.xssf.usermodel.XSSFName;
import org.zkoss.poi.xssf.usermodel.XSSFRelation;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
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
import org.zkoss.zss.model.Worksheet;

/**
 * Implementation of {@link Book} based on XSSFWorkbook.
 * @author henrichen
 *
 */
public class XSSFBookImpl extends XSSFWorkbook implements Book, BookCtrl {
	private final String _bookname;
	private final FormulaEvaluator _evaluator;
	private final WorkbookEvaluator _bookEvaluator;
	private final FunctionMapper _functionMapper;
	private final VariableResolver _variableResolver;
	private RefBook _refBook;
	private BookSeries _bookSeries;
	private int _defaultCharWidth = 7; //TODO: don't know how to calculate this yet per the default font.
	
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
		for(XSSFSheet sheet : this) {
			((SheetCtrl)sheet).initMerged();
		}
		_bookname = bookname;
		FunctionResolver resolver = (FunctionResolver) BookHelper.getLibraryInstance(FunctionResolver.CLASS);
		if (resolver == null) resolver = new DefaultFunctionResolver();
		
		//http://tracker.zkoss.org/browse/ZSS-218
		UDFFinder udff = resolver.getUDFFinder();
		if(udff!=null){
			insertToolPack(0, udff);
		}
		_evaluator = XSSFFormulaEvaluator.create(this, NoCacheClassifier.instance, TolerantUDFFinder.instance); 
		_bookEvaluator = _evaluator.getWorkbookEvaluator(); 
		_bookEvaluator.setDependencyTracker(resolver.getDependencyTracker(this));
		_functionMapper = new JoinFunctionMapper(resolver.getFunctionMapper());
		_variableResolver = new JoinVariableResolver();
	}
	
	/*package*/ WorkbookEvaluator getWorkbookEvaluator() {
		return _bookEvaluator;
	}
	
	/*package*/ RefBook getOrCreateRefBook() {
		if (_refBook == null) {
			_refBook = newRefBook(this);
		}
		return _refBook;
	}
	
	/*package*/ BookSeries getBookSeries() {
		return _bookSeries;
	}
	
	/*package*/ void setBookSeries(BookSeries books) {
		_bookSeries = books;
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
	
	@Override
	public void deletePictureData(PictureData img) {
		getAllPictures().remove(img);
	}
	
	//--Workbook--//
	@Override
	public void removeSheetAt(int index) {
		if (_refBook != null) {
			final String sheetname = getSheetName(index);
			_refBook.removeRefSheet(sheetname);
		}
		super.removeSheetAt(index);
	}

	@Override
	public void setSheetName(int index, String name) {
		final String oldsheetname = getSheetName(index);
		super.setSheetName(index, name);
		if (_refBook != null) {
			_refBook.setSheetName(oldsheetname, name);
		}
		for (XSSFSheet sheet : this) { //scan all sheets to change the possible reference
			((SheetCtrl)sheet).whenRenameSheet(oldsheetname, name);
		}
		ranameSheetInAppXml(oldsheetname, name);
	}

	//rename sheet in app.xml(extended properties)
	private void ranameSheetInAppXml(String oldsheetname, String name) {
	    POIXMLProperties properties = getProperties(); //app.xml and custom.xml 

		CTProperties ext = properties.getExtendedProperties().getUnderlyingProperties();
		CTVectorVariant vectorv = ext.getHeadingPairs();
		CTVector vector = vectorv.getVector();
		CTVariant[] variants = vector.getVariantArray();
		int sheetCount = -1;
		int nameCount = -1;
		for(int j = 0; j < variants.length; ++j) {
			final CTVariant variant = variants[j];
			if ((j & 1) == 0) { //string
				String key = variant.getLpstr();
				if ("Worksheets".equalsIgnoreCase(key)) {
					final CTVariant variant2 = variants[++j];
					sheetCount = variant2.getI4();
				} else if ("Named Ranges".equalsIgnoreCase(key)) {
					final CTVariant variant2 = variants[++j];
					nameCount = variant2.getI4();
				}
			}
		}
		if (sheetCount >= 0 && nameCount >= 0) {
			nameCount += sheetCount;
		}
		CTVectorLpstr vectorv2 = ext.getTitlesOfParts();
		CTVector vector2 = vectorv2.getVector();
		String[] lpstrs = vector2.getLpstrArray();
		int j = 0;
		for(; j < sheetCount; ++j) {
			final String sname = lpstrs[j];
			if (oldsheetname.equals(sname)) {
				vector2.setLpstrArray(j, name);
				j = sheetCount;
				break;
			}
		}
		final String o = SheetNameFormatter.format(oldsheetname);
		final String n = SheetNameFormatter.format(name);
		for(; j < nameCount; ++j) {
			final String refname = lpstrs[j];
			final String newrefname = refname.replace(o+"!", n+"!");
			if (!newrefname.equals(refname)) {
				vector2.setLpstrArray(j, newrefname);
			}
		}
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
    public XSSFName getBuiltInName(String builtInCode, int sheetNumber) {
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
	public Worksheet getWorksheet(String name) {
		return (Worksheet) getSheet(name);
	}
	
	@Override
	public boolean isDate1904() {
		return super.isDate1904();
	}

	@Override
	public Worksheet getWorksheetAt(int index) {
		return (Worksheet) getSheetAt(index);
	}

    
	//--BookCtrl--//
	@Override
	public RefBook newRefBook(Book book) {
		return getBookCtrl().newRefBook(book);
	}

	@Override
	public Object nextSheetId() {
		return getBookCtrl().nextSheetId();
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
}
