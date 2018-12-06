/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

import org.zkoss.lang.Objects;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.util.logging.Log;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.model.EventQueueModelEventListener;
import org.zkoss.zss.model.InvalidModelOpException;
import org.zkoss.zss.model.ModelEvent;
import org.zkoss.zss.model.ModelEventListener;
import org.zkoss.zss.model.ModelEvents;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColor;
import org.zkoss.zss.model.SColumnArray;
import org.zkoss.zss.model.SExtraStyle;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SName;
import org.zkoss.zss.model.SPicture;
import org.zkoss.zss.model.SPictureData;
import org.zkoss.zss.model.SRow;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SNamedStyle;
import org.zkoss.zss.model.STable;
import org.zkoss.zss.model.STableColumn;
import org.zkoss.zss.model.STableStyle;
import org.zkoss.zss.model.STableStyleElem;
import org.zkoss.zss.model.SheetRegion;
import org.zkoss.zss.model.impl.sys.DependencyTableAdv;
import org.zkoss.zss.model.impl.sys.formula.ParsingBook;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.dependency.Ref.RefType;
import org.zkoss.zss.model.sys.formula.EvaluationContributor;
import org.zkoss.zss.model.sys.formula.FormulaClearContext;
import org.zkoss.zss.model.util.CellStyleMatcher;
import org.zkoss.zss.model.util.FontMatcher;
import org.zkoss.zss.model.util.Strings;
import org.zkoss.zss.model.util.Validations;
import org.zkoss.zss.range.impl.NotifyChangeHelper;
import org.zkoss.zss.range.impl.StyleUtil;

/**
 * @author dennis
 * @since 3.5.0
 */
public class BookImpl extends AbstractBookAdv{
	private static final long serialVersionUID = 1L;
	
	private static final Log _logger = Log.lookup(BookImpl.class);

	private static final String SYNC_MERGE_FIRED = "_ZSS_MER";
	
	private final String _bookName;
	
	private String _shareScope;
	
	private SBookSeries _bookSeries;
	
	private final List<AbstractSheetAdv> _sheets = new ArrayList<AbstractSheetAdv>();
	private List<AbstractNameAdv> _names;
	
	private final List<SCellStyle> _cellStyles = new ArrayList<SCellStyle>(); 
	private final Map<String, SNamedStyle> _namedStyles = new HashMap<String, SNamedStyle>(); //ZSS-854
	private final List<SCellStyle> _defaultCellStyles = new ArrayList<SCellStyle>(); //ZSS-854
	private final List<AbstractFontAdv> _fonts = new ArrayList<AbstractFontAdv>();
	private AbstractFontAdv _defaultFont;
	private final HashMap<AbstractColorAdv,AbstractColorAdv> _colors = new LinkedHashMap<AbstractColorAdv,AbstractColorAdv>();
	
	private final static Random _random = new Random(System.currentTimeMillis());
	private final static AtomicInteger _bookCount = new AtomicInteger();
	private final String _bookId;
	
	private final HashMap<String,AtomicInteger> _objIdCounter = new HashMap<String,AtomicInteger>();
	private final int _maxRowSize = SpreadsheetVersion.EXCEL2007.getMaxRows();
	private final int _maxColumnSize = SpreadsheetVersion.EXCEL2007.getMaxColumns();
	
	private EventListenerAdaptor _listeners;
	private EventListenerAdaptor _queueListeners;
	
	private HashMap<String,Object> _attributes;
	
	private EvaluationContributor _evalContributor;
	
	private ArrayList<SPictureData> _picDatas; //since 3.6.0
	
	private boolean _dirty = false;
	
	//ZSS-855
	private HashMap<String, STable> _tables; //since 3.8.0
	
	//ZSS-1140
	private List<SExtraStyle> _extraStyles = new ArrayList<SExtraStyle>(); // since 3.8.2
	
	//ZSS-992
	private LinkedHashMap<String, STableStyle> _tableStyles = new LinkedHashMap<String, STableStyle>(); // since 3.8.3
	//ZSS-992
	private String _defaultPivotStyle; //default pivot style name; since 3.8.3
	//ZSS-992
	private String _defaultTableStyle; //default table style name; since 3.8.3
	//ZSS-1283
	private boolean _inPostProcessing; // whether the book is in post import processing cycle
	/**
	 * the sheet which is destroying now.
	 */
	/*package*/ final static ThreadLocal<SSheet> destroyingSheet = new ThreadLocal<SSheet>(); 
	
	public BookImpl(String bookName){
		Validations.argNotNull(bookName);
		this._bookName = bookName;
		_bookSeries = new SimpleBookSeriesImpl(this);
		_fonts.add(_defaultFont = new FontImpl());
		initDefaultCellStyles();
		_colors.put(ColorImpl.WHITE,ColorImpl.WHITE);
		_colors.put(ColorImpl.BLACK,ColorImpl.BLACK);
		_colors.put(ColorImpl.RED,ColorImpl.RED);
		_colors.put(ColorImpl.GREEN,ColorImpl.GREEN);
		_colors.put(ColorImpl.BLUE,ColorImpl.BLUE);
		
		_bookId = ((char)('a'+_random.nextInt(26))) + Long.toString(/*System.currentTimeMillis()+*/_bookCount.getAndIncrement(), Character.MAX_RADIX) ;
		_tables = new HashMap<String, STable>(0);
	}
	
	public void initDefaultCellStyles() {
		AbstractCellStyleAdv defaultCellStyle = new CellStyleImpl(_defaultFont);
		_cellStyles.add(defaultCellStyle); //ZSS-854
		_defaultCellStyles.add(defaultCellStyle); //ZSS-854
	}
	
	@Override
	public SBookSeries getBookSeries(){
		return _bookSeries;
	}
	
	@Override
	public String getBookName(){
		return _bookName;
	}
	
	@Override
	public SSheet getSheet(int i){
		return _sheets.get(i);
	}
	
	@Override
	public int getNumOfSheet(){
		return _sheets.size();
	}
	
	@Override
	public SSheet getSheetByName(String name){
		for(SSheet sheet:_sheets){
			if(sheet.getSheetName().equalsIgnoreCase(name)){
				return sheet;
			}
		}
		return null;
	}
	
	@Override
	public SSheet getSheetById(String id){
		for(SSheet sheet:_sheets){
			if(sheet.getId().equals(id)){
				return sheet;
			}
		}
		return null;
	}
	
	protected void checkOwnership(SSheet sheet){
		if(!_sheets.contains(sheet)){
			throw new IllegalStateException("doesn't has ownership "+ sheet);
		}
	}
	protected void checkOwnership(SName name){
		if(_names==null || !_names.contains(name)){
			throw new IllegalStateException("doesn't has ownership "+ name);
		}
	}

	@Override
	public void sendModelEvent(ModelEvent event){
		if(_listeners!=null){
			_listeners.sendModelEvent(event);
		}
		
		if(_queueListeners!=null) {
			// System thread doesn't have execution so that it will throw IllegalStateException
			// e.g. Background Thread created by Executor
			final AbstractSheetAdv sheet = (AbstractSheetAdv)event.getSheet(); 
			if (Executions.getCurrent() != null) {
				//ZSS-1168, 20151223, henrichen:
				// When {@link MergeHelper#merge()} or {@link MergeHelper#unmerge()}, 
				// sheet#mergeOutOfSync is set to true.
				final String eventName = event.getName();
				if ((ModelEvents.ON_MERGE_ADD.equals(eventName) 
				|| ModelEvents.ON_MERGE_DELETE.equals(eventName))
				&& sheet != null && sheet.getMergeOutOfSync() == 1) {
					sheet.setMergeOutOfSync(0);
				} else if (!ModelEvents.ON_MERGE_SYNC.equals(eventName)
					&& !ModelEvents.ON_SHEET_DELETE.equals(eventName) //ZSS-1200: sheet is to be deleted, no need to handle the sheet
					&& sheet != null && sheet.getMergeOutOfSync() == 2) {
					// notify all associated Spreadsheets to clear the merge cache first 
					sheet.setMergeOutOfSync(0);
					//ZSS-1204: notifyMergeSync should be called only once in an execution
					Boolean syncMerged = (Boolean) Executions.getCurrent().getAttribute(SYNC_MERGE_FIRED, false);
					if (syncMerged == null) {
						Executions.getCurrent().setAttribute(SYNC_MERGE_FIRED, Boolean.TRUE, false);
						new NotifyChangeHelper().notifyMergeSync(new SheetRegion((SSheet)sheet, 1, 1));
					}
				}
				_queueListeners.sendModelEvent(event);
			} else if (sheet != null && sheet.getMergeOutOfSync() == 1) { 
				//ZSS-1168: in long operation event queue and merge changed
				sheet.setMergeOutOfSync(2);
			}
		}
		
		if(!ModelEvents.isCustomEvent(event)) {
			if(!_dirty) {
				_dirty = true;
				// ZSS-942, By Jerry 2015/3/5
				// ATTENTION: ModelEvents.ON_MODEL_DIRTY_CHANGE is a custom event.
				// Dirty change is a special case for calling sendModelEvent inside model.
				// In normal model event case, we should do it in Range level.
				sendModelEvent(ModelEvents.createModelEvent(ModelEvents.ON_MODEL_DIRTY_CHANGE, event.getBook(), event.getSheet(),
						ModelEvents.createDataMap(ModelEvents.PARAM_CUSTOM_DATA, _dirty)));
			}
		}
	}
	
	@Override
	public SSheet createSheet(String name) {
		return createSheet(name,null);
	}
	
	@Override
	String nextObjId(String type){
		StringBuilder sb = new StringBuilder(_bookId);
		sb.append("_").append(type).append("_");
		AtomicInteger i = _objIdCounter.get(type);
		if(i==null){
			_objIdCounter.put(type, i = new AtomicInteger(0));
		}
		sb.append(i.getAndIncrement());
		return sb.toString();
	}
	
	@Override
	public SSheet createSheet(String name, SSheet src) {
		checkLegalSheetName(name);
		
		// ZSS-1183: Support createSheet from a source sheet no matter the sheet is
		// in the same book or not. However, still have to check if the src
		// sheet is an orphan sheet.		
		if(src!=null)
			((SheetImpl)src).checkOrphan();
		

		AbstractSheetAdv sheet = new SheetImpl(this,nextObjId("sheet"));
		sheet.setSheetName(name);
		_sheets.add(sheet);
		
		if(src instanceof AbstractSheetAdv){
			((AbstractSheetAdv)src).copyTo(sheet);
		}
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));

		//ZSS-1283
		if (!this.isPostProcessing()) {
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(),new RefImpl(sheet, -1));
		}

		return sheet;
	}

	protected Ref getRef() {
		return new RefImpl(this);
	}

	@Override
	public void setSheetName(SSheet sheet, String newname) {
		checkLegalSheetName(newname);
		checkOwnership(sheet);
		
		int index = getSheetIndex(sheet);
		String oldname = sheet.getSheetName();
		((AbstractSheetAdv)sheet).setSheetName(newname);
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		
		//ZSS-1283
		if (!this.isPostProcessing()) {
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(),new RefImpl(this.getBookName(),newname, index));//to clear the cache of formula that has unexisted name
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(),new RefImpl(this.getBookName(),oldname, index));
		}
		
		renameSheetFormula(oldname,newname,index);
	}
	
	private void renameSheetFormula(String oldName, String newName, int index){
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)getBookSeries();
		DependencyTable dt = bs.getDependencyTable();
		Set<Ref> dependents = dt.getDirectDependents(new RefImpl(getBookName(),oldName,index));
		if(dependents.size()>0){
			
			//clear the dependents dependency before rename it's sheet name
			for(Ref dependent:dependents){
				dt.clearDependents(dependent);
			}
			
			//rebuild the the formula by tuner
			FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
			tuner.renameSheet(this,oldName,newName,dependents);
		}
		
		//ZSS-1137
		// rename sheetScope of the SName
		for (SName name : getNames()) {
			if (oldName.equalsIgnoreCase(name.getApplyToSheetName())) {
				name.setApplyToSheetName(newName);
			}
		}
	}

	private void checkLegalSheetName(String name) {
		if(Strings.isBlank(name)){
			throw new InvalidModelOpException("sheet name '"+name+"' is not legal");
		}
		if(getSheetByName(name)!=null){
			throw new InvalidModelOpException("sheet name '"+name+"' is duplicated");
		}
	}
	
	private void checkLegalNameName(String name,String sheetName) {
		if(Strings.isBlank(name)){
			throw new InvalidModelOpException("name '"+name+"' is not legal");
		}
		//ZSS-966
		if (getTable(name) != null) {
			throw new InvalidModelOpException("name '"+name+"' is duplicated with Table name");
		}
		if(getNameByName(name,sheetName)!=null){ //must be unique in the scope
			throw new InvalidModelOpException("name '"+name+"' "+(sheetName==null?"":" in '"+sheetName+"'")+" is duplicated");
		}
		if(sheetName!=null && getSheetByName(sheetName)==null){
			throw new InvalidModelOpException("no such sheet "+sheetName);
		}
		//ZSS-660: valid name
		//@see  http://office.microsoft.com/en-us/excel-help/define-and-use-names-in-formulas-HA010147120.aspx
		//length must less than or equals to 255
		if (name.length() > 255) {
			throw new InvalidModelOpException("name '"+name+"' is not legal: cannot exceed 255 characters");
		}
		
		//1st character must be a letter, underscore, or backslash
		char c1 = name.charAt(0);
		if (!Character.isLetter(c1) && c1 != '_' && c1 != '\\') {
			throw new InvalidModelOpException("name '"+name+"' is not legal: first character must be a letter, an underscore, or a backslash");
		}
		
		boolean invalid = c1 == '_' || c1 == '\\' || c1 == '?' || c1 == '.'; //impossible be a valid cell reference
		int colIndex = invalid ? -2 : Character.getNumericValue(c1) - 9;
		if (!invalid) {
			invalid = colIndex < 0;
		}
		int rowIndex = -1;
		//remaining characters must be letters, digits, periods, or underscores.
		for (int j = 1, len = name.length(); j < len; ++j) {
			char ch = name.charAt(j);
			if (Character.isLetter(ch)) { //analyze colIndex
				if (invalid) continue;  
				if (rowIndex >= 0) { //letter -> digit -> letter
					invalid = true;
					continue;
				}
				int c = Character.getNumericValue(ch) - 9;
				if (c < 0) {
					invalid = true;
					continue;
				}
				colIndex = colIndex * 26 + c;
			} else if (Character.isDigit(ch)) { //analyze rowIndex
				if (invalid) continue; 
				if (rowIndex < 0) {
					rowIndex = Character.getNumericValue(ch);
				} else {
					rowIndex = rowIndex * 10 + Character.getNumericValue(ch);
				}
			} else if (ch != '.' && ch != '_' && ch != '?' && ch != '\\') {
				throw new InvalidModelOpException("name '"+name+"' is not legal: the character '"+ ch+ "' at index "+ j + " must be a letter, a digit, an underscore, a period, a question mark, or a backslash");
			} else { //ZSS-792
				invalid = true; // '.' or '-' or '?' or '\', impossible to be a valid cell reference
			}
		}
		
		//cannot be a valid cell reference address
		if (!invalid && colIndex >= 0 && colIndex <= getMaxColumnSize() && rowIndex >= 0 && rowIndex < getMaxRowSize()) {
			throw new InvalidModelOpException("name '"+name+"' is not legal: cannot be a cell reference");
		}
			
		//cannot be 'C' or 'R'
		if (name.equalsIgnoreCase("C") || name.equalsIgnoreCase("R")) {
			throw new InvalidModelOpException("name '"+name+"' is not legal: cannot be 'C', 'c', 'R', or 'r'");
		}
	}

	@Override
	public void deleteSheet(SSheet sheet) {
		checkOwnership(sheet);
		
		final String bookName = sheet.getBook().getBookName();
		
		destroyingSheet.set(sheet);
		try{
			((AbstractSheetAdv)sheet).destroy();
		}finally{
			destroyingSheet.set(null);
		}
		String oldName = sheet.getSheetName();
		int index = _sheets.indexOf(sheet);
		_sheets.remove(index);
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		
//		sendModelInternalEvent(ModelInternalEvents.createModelInternalEvent(ModelInternalEvents.ON_SHEET_DELETED, 
//				this,ModelInternalEvents.createDataMap(ModelInternalEvents.PARAM_SHEET_OLD_INDEX, index)));
		
		//ZSS-1283
		if (!this.isPostProcessing()) {
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(),new RefImpl(this.getBookName(),sheet.getSheetName(), index));
		}
		
		renameSheetFormula(oldName, null, index);
		
		//ZSS-815
		// adjust sheet index
		adjustSheetIndex(bookName, index);
	}

	//ZSS-815
	private void adjustSheetIndex(String bookName, int index) {
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)getBookSeries();
		DependencyTableAdv dt = (DependencyTableAdv) bs.getDependencyTable();
		dt.adjustSheetIndex(bookName, index, -1);
	}
		
	@Override
	public void moveSheetTo(SSheet sheet, int index) {
		checkOwnership(sheet);
		if(index<0|| index>=_sheets.size()){
			throw new InvalidModelOpException("new position out of bound "+_sheets.size() +"<>" +index);
		}
		int oldindex = _sheets.indexOf(sheet);
		if(oldindex==index){
			return;
		}
		
		//ZSS-820
		reorderSheetFormula(getSheet(oldindex).getSheetName(), oldindex, index);
		_sheets.remove(oldindex);
		_sheets.add(index, (AbstractSheetAdv)sheet);
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));

		//ZSS-1283
		if (!this.isPostProcessing()) {		
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(),new RefImpl(this.getBookName(),sheet.getSheetName(), index));
			//ZSS-1049: should consider formulas that referred to the old index 
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(),new RefImpl(this.getBookName(),sheet.getSheetName(), oldindex));
		}
		// adjust sheet index
		moveSheetIndex(getBookName(), oldindex, index);
	}
	//ZSS-820
	private void moveSheetIndex(String bookName, int oldIndex, int newIndex) {
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)getBookSeries();
		DependencyTableAdv dt = (DependencyTableAdv) bs.getDependencyTable();
		dt.moveSheetIndex(bookName, oldIndex, newIndex);
	}

	public void dump(StringBuilder builder) {
		for(AbstractSheetAdv sheet:_sheets){
			if(sheet instanceof SheetImpl){
				((SheetImpl)sheet).dump(builder);
			}else{
				builder.append("\n").append(sheet);
			}
		}
	}

	@Override
	public SCellStyle getDefaultCellStyle() {
		return getDefaultCellStyle(0);
	}

	//ZSS-854
	@Override
	public SCellStyle getDefaultCellStyle(int index) {
		return _defaultCellStyles.get(index);
	}

	@Override
	public void setDefaultCellStyle(SCellStyle cellStyle) {
		if (cellStyle == null) return;
		AbstractCellStyleAdv defaultCellStyle = (AbstractCellStyleAdv) cellStyle;
		_defaultCellStyles.set(0, defaultCellStyle);
		_cellStyles.set(0, defaultCellStyle);
	}

	@Override
	public SCellStyle createCellStyle(boolean inStyleTable) {
		return createCellStyle(null,inStyleTable);
	}

	@Override
	public SCellStyle createCellStyle(SCellStyle src,boolean inStyleTable) {
		if(src!=null){
			Validations.argInstance(src, AbstractCellStyleAdv.class);
		}
		AbstractCellStyleAdv style = new CellStyleImpl(_defaultFont);
		if(src!=null){
			style.copyFrom(src);
		}
		
		if(inStyleTable){
			_cellStyles.add(style);
		}
		
		return style;
	}
	
	@Override
	public SCellStyle searchCellStyle(CellStyleMatcher matcher) {
		for(SCellStyle style:_cellStyles){
			if(matcher.match(style)){
				return style;
			}
		}
		return null;
	}

	//ZSS-1183
	@Override
	public SExtraStyle searchExtraStyle(CellStyleMatcher matcher) {
		for(SExtraStyle style:_extraStyles){
			if(matcher.match(style)){
				return style;
			}
		}
		return null;
	}
	
	@Override
	public SFont getDefaultFont() {
		return _defaultFont;
	}

	@Override
	public SFont createFont(boolean inFontTable) {
		return createFont(null,inFontTable);
	}

	@Override
	public SFont createFont(SFont src,boolean inFontTable) {
		if(src!=null){
			Validations.argInstance(src, AbstractFontAdv.class);
		}
		AbstractFontAdv font = new FontImpl();
		if(src!=null){
			font.copyFrom(src);
		}
		
		if(inFontTable){
			_fonts.add(font);
		}
		
		return font;
	}
	
	@Override
	public SFont searchFont(FontMatcher matcher) {
		for(SFont font:_fonts){
			if(matcher.match(font)){
				return font;
			}
		}
		return null;
	}
	
	@Override
	public int getMaxRowSize() {
		return _maxRowSize;
	}

	@Override
	public int getMaxColumnSize() {
		return _maxColumnSize;
	}

	@Override
	public void optimizeCellStyle() {
		HashMap<String,SCellStyle> stylePool = new LinkedHashMap<String,SCellStyle>();
		_cellStyles.clear();
		_fonts.clear();
		
		SCellStyle defaultStyle = getDefaultCellStyle();
		SFont defaultFont = getDefaultFont();
		stylePool.put(((AbstractCellStyleAdv)defaultStyle).getStyleKey(), defaultStyle);
		
		for(SSheet sheet:_sheets){
			Iterator<SRow> rowIter = sheet.getRowIterator(); 
			while(rowIter.hasNext()){
				SRow row = rowIter.next();
				
				row.setCellStyle(hitStyle(defaultStyle,row.getCellStyle(),stylePool));
				Iterator<SCell> cellIter = sheet.getCellIterator(row.getIndex());
				while(cellIter.hasNext()){
					SCell cell = cellIter.next();
					cell.setCellStyle(hitStyle(defaultStyle,cell.getCellStyle(),stylePool));
				}
			}
			Iterator<SColumnArray> colIter = sheet.getColumnArrayIterator();
			while(colIter.hasNext()){
				SColumnArray colarr = colIter.next();
				colarr.setCellStyle(hitStyle(defaultStyle,colarr.getCellStyle(),stylePool));
			}
		}
		
		_cellStyles.addAll((Collection)stylePool.values());
		String key;
		HashMap<String,SFont> fontPool = new LinkedHashMap<String,SFont>();
		
		fontPool.put(((AbstractFontAdv)defaultFont).getStyleKey(), defaultFont);
		for(SCellStyle style:_cellStyles){
			SFont font = style.getFont();
			key = ((AbstractFontAdv)font).getStyleKey();
			if(fontPool.get(key)==null){
				fontPool.put(key, font);
			}
		}
		
		_fonts.addAll((Collection)fontPool.values());
		
		_colors.clear();//color is immutable, just clear it.
	}
	
	
	@SuppressWarnings("unchecked")
	public List<SCellStyle> getCellStyleTable(){
		return Collections.unmodifiableList((List)_cellStyles);
	}
	@SuppressWarnings("unchecked")
	public List<SFont> getFontTable(){
		return Collections.unmodifiableList((List)_fonts);
	}
	
	private SCellStyle hitStyle(SCellStyle defaultStyle,SCellStyle currSytle,
			HashMap<String, SCellStyle> stylePool) {
		String key;
		SCellStyle hit;
		if(currSytle==defaultStyle){//quick case for most cell use default style
			return defaultStyle;
		}else{
			key = ((AbstractCellStyleAdv)currSytle).getStyleKey();
			hit = stylePool.get(key);
			if(hit==null){
				stylePool.put(key, hit = currSytle);
			}
		}
		return hit;
	}

	@Override
	public void addEventListener(ModelEventListener listener){
		if(listener instanceof EventQueueModelEventListener){
			if(_queueListeners==null){
				String scope = getShareScope();
				if(scope==null){
					scope = "desktop";//default desktop
				}
				_queueListeners = new EventQueueListenerAdaptor(scope,getId());
			}
			_queueListeners.addEventListener(listener);
		}else{
			if(_listeners==null){
				_listeners = new DirectEventListenerAdaptor();
			}
			_listeners.addEventListener(listener);
		}
	}
	@Override
	public void removeEventListener(ModelEventListener listener){
		if(listener instanceof EventQueueModelEventListener && _queueListeners!=null){
			_queueListeners.removeEventListener(listener);
			if(_queueListeners.size()==0){
				_queueListeners = null;//clean up, so user can change share-scope then.
			}
		}else if(_listeners!=null){
			_listeners.removeEventListener(listener);
		}
	}

	@Override
	public Object getAttribute(String name) {
		return _attributes==null?null:_attributes.get(name);
	}

	@Override
	public Object setAttribute(String name, Object value) {
		if(_attributes==null){
			_attributes = new HashMap<String, Object>();
		}
		return _attributes.put(name, value);
	}

	@Override
	public Map<String, Object> getAttributes() {
		return _attributes==null?Collections.EMPTY_MAP:Collections.unmodifiableMap(_attributes);
	}

	@Override
	public SColor createColor(byte r, byte g, byte b) {
		AbstractColorAdv newcolor = new ColorImpl(r,g,b);
		AbstractColorAdv color = _colors.get(newcolor);//reuse the existed color object
		if(color==null){
			_colors.put(newcolor, color = newcolor);
		}
		return color;
	}

	@Override
	public SColor createColor(String htmlColor) {
		AbstractColorAdv newcolor = new ColorImpl(htmlColor);
		AbstractColorAdv color = _colors.get(newcolor);//reuse the existed color object
		if(color==null){
			_colors.put(newcolor, color = newcolor);
		}
		return color;
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public List<SSheet> getSheets() {
		return Collections.unmodifiableList((List)_sheets);
	}
	@Override
	public SName createName(String namename) {
		return createName(namename,null);
	}
	@Override
	public SName createName(String namename,String sheetName) {
		checkLegalNameName(namename,sheetName);

		AbstractNameAdv name = new NameImpl(this,nextObjId("name"),namename,sheetName);
		
		if(_names==null){
			_names = new ArrayList<AbstractNameAdv>();
		}
		
		_names.add(name);
		return name;
	}

	@Override
	public void setNameName(SName name, String newname) {
		setNameName(name,newname,null);
	}
	public void setNameName(SName name, String newname, String sheetName) {
		checkLegalNameName(newname,sheetName);
		checkOwnership(name);
		
		//create formula cache for name, currently, we can just clear all of book.
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		
		//notify the (old) name is change before update name
		//ZSS-1283
		if (!this.isPostProcessing()) {
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(),new NameRefImpl((AbstractNameAdv)name));
		}
		
		final String oldName = name.getName(); // ZSS-661
		
		//ZSS-966: notify the (old) table name is change before update name
		//ZSS-1283
		if (name instanceof TableNameImpl && !this.isPostProcessing()) {
			ModelUpdateUtil.handlePrecedentUpdate(getBookSeries(), new TablePrecedentRefImpl(this.getBookName(), oldName));
		}
		
		((AbstractNameAdv)name).setName(newname,sheetName); //will change Table's name if the name is a TableName
		//don't need to notify new name precedent update, since Name handle it itself
		
		//Rename formula that contains this name
		renameNameFormula(name, oldName, newname, sheetName); // ZSS-661
		
		//Rename formula that contains this table name
		if (name instanceof TableNameImpl) {
			//ZSS-966: reput Table; must do this first or renameTableName() 
			//cannot find the correct table
			STable tb = removeTable(oldName);
			if (tb != null)
				addTable(tb);

			renameTableNameFormula(name, oldName, newname);
		}
	}
	
	private void renameNameFormula(SName name, String oldName, String newName, String sheetName) {
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)getBookSeries();
		FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
		DependencyTable dt = bs.getDependencyTable();
		final 
		Ref ref = new NameRefImpl(name.getBook().getBookName(),name.getApplyToSheetName(), oldName); //old name
		Set<Ref> dependents = dt.getDirectDependents(ref);
		if(dependents.size()>0){
			//clear the dependents dependency before rename it's Name name
			for(Ref dependent:dependents){
				dt.clearDependents(dependent);
			}
			
			int sheetIndex = sheetName == null ? -1 : getSheetIndex(sheetName);
			//rebuild the the formula by tuner
			tuner.renameName(this, oldName, newName, dependents, sheetIndex);
		}
	}

	@Override
	public void deleteName(SName name) {
		checkOwnership(name);
		
		((AbstractNameAdv)name).destroy();
		
		int index = _names.indexOf(name);
		_names.remove(index);
		
//		sendEvent(ModelEvents.ON_NAME_DELETED, 
//				ModelEvents.PARAM_NAME, sheet,
//				ModelEvents.PARAM_SHEET_OLD_INDEX, index);
	}

	@Override
	public int getNumOfName() {
		return _names==null?0:_names.size();
	}

	@Override
	public SName getName(int idx) {
		if(_names==null){
			throw new ArrayIndexOutOfBoundsException(idx);
		}
		return _names.get(idx);
	}

	@Override
	public SName getNameByName(String namename) {
		return getNameByName(namename,null);
	}
	public SName getNameByName(String namename, String sheetName) {
		if(_names==null || (sheetName != null && getSheetByName(sheetName) == null)) //ZSS-1137
			return null;
		for(SName name:_names){
			//ZSS-436
			final String scopeSheetName = name.getApplyToSheetName();
			if ((sheetName == scopeSheetName || (sheetName != null && sheetName.equalsIgnoreCase(scopeSheetName)))
					&& name.getName().equalsIgnoreCase(namename)){
				return name;
			}
		}
		return null;
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public List<SName> getNames() {
		return _names==null?Collections.EMPTY_LIST:Collections.unmodifiableList((List)_names);
	}

	@Override
	public int getSheetIndex(SSheet sheet) {
		return _sheets.indexOf(sheet);
	}
	
	@Override
	public int getSheetIndex(String sheetName) {
		int i=0;
		for(SSheet sheet:_sheets){
			if(sheet.getSheetName().equals(sheetName)){
				return i;
			}
			i++;
		}
		return -1;
	}

	@Override
	public void setShareScope(String scope) {
		if(!Objects.equals(this._shareScope,scope)){
			
			if("disable".equals(scope)){
				if(_listeners!=null){
					_listeners.clear();
				}
				if(_queueListeners!=null){
					_queueListeners.clear();
				}
				return;
			}
			
			if(_queueListeners!=null && _queueListeners.size()>0){
				throw new IllegalStateException("can't change share scope after registed any queue model event listener");
			}
			
			this._shareScope = scope;
		}
	}

	@Override
	public String getShareScope() {
		return _shareScope;
	}

	@Override
	void setBookSeries(SBookSeries bookSeries) {
		this._bookSeries = bookSeries;
	}

	@Override
	public EvaluationContributor getEvaluationContributor() {
		return _evalContributor;
	}

	@Override
	public void setEvaluationContributor(EvaluationContributor contributor) {
		this._evalContributor = contributor;
	}

	@Override
	public int getMaxRowIndex() {
		return getMaxRowSize()-1;
	}

	@Override
	public int getMaxColumnIndex() {
		return getMaxColumnSize()-1;
	}

	@Override 
	public String getId(){
		return _bookId; 
	}
	
	@Override
	public SPictureData addPictureData(SPicture.Format format, byte[] data) {
		if (_picDatas == null) {
			_picDatas = new ArrayList<SPictureData>(4);
		}
		int index = _picDatas.size();
		SPictureData picData = new PictureDataImpl(index, format, data);
		_picDatas.add(picData);
		return picData;
	}

	@Override
	public SPictureData getPictureData(int index) {
		if (index < 0 || _picDatas == null || index >= _picDatas.size())
			return null;
		return _picDatas.get(index);
	}

	@Override
	public Collection<SPictureData> getPicturesDatas() {
		if (_picDatas == null) return Collections.emptyList();
		
		final List<SPictureData> list = new ArrayList<SPictureData>(_picDatas.size());
		for (SPictureData picData : _picDatas) {
			if (picData != null) {
				list.add(picData);
			}
		}
		return list;
	}

	//ZSS-820
	private void reorderSheetFormula(String sheetName, int oldIndex, int newIndex) {
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)getBookSeries();
		DependencyTable dt = bs.getDependencyTable();
		Set<Ref> dependents = dt.getDirectDependents(new RefImpl(getBookName(),sheetName, oldIndex));
		if(dependents.size()>0){
			final String bookName = getBookName();
			final int low = oldIndex < newIndex ? oldIndex : newIndex;
			final int high = oldIndex < newIndex ? newIndex : oldIndex;

			Set<String> bookNames = new HashSet<String>();
			//filter out dependents that does not need to do reorder
			for(final Iterator<Ref> it = dependents.iterator(); it.hasNext();) {
				Ref dependent = it.next();
				Set<Ref> precedents = ((DependencyTableAdv)dt).getDirectPrecedents(dependent);
				boolean candidate = false;
				for (Ref p : precedents) {
					if (p.getType() != RefType.AREA && p.getType() != RefType.CELL)
						continue;
					final String bookName0 = p.getBookName();
					final String sheet1 = p.getSheetName();
					final String sheet2 = p.getLastSheetName();
					
					final int low0 = getSheetIndex(sheet1);
					int high0 = getSheetIndex(sheet2);
					if (high0 < 0) high0 = low0;
					if (low0 == high0) continue; // single sheet, as is.
							
					// no intersection; as is.
					if (high0 < low || low0 > high) continue;

					if (low0 == oldIndex) {
						if (low0 != high0 && newIndex >= high0) { //2. move beyond original high end
							//must change low end sheet name! (_map & _remap must be remapped)
							candidate = true;
							
							// adjust extern sheet name
							if (!bookNames.contains(bookName0)) {
								ParsingBook parsingBook = new ParsingBook(bs.getBook(bookName0));
								parsingBook.reorderSheet(bookName, oldIndex, newIndex);
								bookNames.add(bookName0);
							}
							break;
						}
					}
					
					if (high0 == oldIndex) {
						if (low0 != high0 && newIndex <= low0) { //4. move beyond original low end
							// high0 index not change but sheet name changed
							candidate = true;
							
							// adjust extern sheet name
							if (!bookNames.contains(bookName0)) {
								ParsingBook parsingBook = new ParsingBook(bs.getBook(bookName0));
								parsingBook.reorderSheet(bookName, oldIndex, newIndex);
								bookNames.add(bookName0);
							}
							break;
						}
					}
				}
				if (!candidate) {
					it.remove();
				}
			}

			if (!dependents.isEmpty()) {
				for (Ref dependent : dependents) {
					dt.clearDependents(dependent);
				}
			
				//rebuild the the formula by tuner
				FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
				tuner.reorderSheet(this, oldIndex, newIndex, dependents);
			}
		}
	}
	
	//ZSS-854
	public SNamedStyle getNamedStyle(String name) {
		return _namedStyles.get(name);
	}

	//ZSS-854
	@Override
	public int addDefaultCellStyle(SCellStyle cellStyle) {
		_defaultCellStyles.add((AbstractCellStyleAdv)cellStyle);
		return _defaultCellStyles.size() - 1;
	}
	
	//ZSS-854
	@Override
	public Collection<SCellStyle> getDefaultCellStyles() {
		return _defaultCellStyles;
	}

	//ZSS-854
	@Override
	public void addNamedCellstyle(SNamedStyle namedStyle) {
		_namedStyles.put(namedStyle.getName(), namedStyle);
	}

	//ZSS-854
	@Override
	public Collection<SNamedStyle> getNamedStyles() {
		return _namedStyles.values();
	}

	//ZSS-854
	@Override
	public void clearDefaultCellStyles() {
		_cellStyles.clear();
		_defaultCellStyles.clear();
	}

	//ZSS-854
	@Override
	public void clearNamedStyles() {
		_namedStyles.clear();		
	}

	//ZSS-923
	@Override
	public boolean isDirty() {
		return _dirty;
	}

	//ZSS-923
	@Override
	public void setDirty(boolean dirty) {
		_dirty = dirty;
	}
	
	//ZSS-855
	@Override
	public SName createTableName(STable table) {
		final String namename = table.getName();
		checkLegalNameName(namename, null);

		AbstractNameAdv name = new TableNameImpl(this,table,nextObjId("name"),namename);
		
		if(_names==null){
			_names = new ArrayList<AbstractNameAdv>();
		}
		
		_names.add(name);
		return name;
	}
	
	//ZSS-855
	@Override
	public void addTable(STable table) {
		_tables.put(table.getName().toUpperCase(), table);
	}

	//ZSS-855
	@Override
	public STable getTable(String name) {
		return _tables.get(name.toUpperCase());
	}
	
	//ZSS-855
	@Override
	public STable removeTable(String name) {
		final STable tb = _tables.remove(name.toUpperCase());
		//ZSS-988: should consider table filter
		if (tb != null) {
			((AbstractTableAdv)tb).refreshFilter();
		}
		return tb; 
	}
	
	//ZSS-966
	private void renameTableNameFormula(SName name, String oldName, String newName) {
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)getBookSeries();
		FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
		DependencyTable dt = bs.getDependencyTable();
		final Ref ref = new TablePrecedentRefImpl(name.getBook().getBookName(), oldName); //old name
		Set<Ref> dependents = dt.getDirectDependents(ref);
		if(dependents.size()>0){
			//clear the dependents dependency before rename it's Name name
			for(Ref dependent:dependents){
				dt.clearDependents(dependent);
			}
			
			//rebuild the the formula by tuner
			tuner.renameTableName(this, oldName, newName, dependents);
		}
	}
	
	//ZSS-967
	public String setTableColumnName(STable table, String oldName, String newName) {
		if (Objects.equals(oldName, newName)) return newName;
		
		// locate the STableColumn of oldName
		List<STableColumn> tbCols = table.getColumns();
		STableColumn tbCol = null;
		STableColumn tbColDup = null;
		Set<String> set = new HashSet<String>(tbCols.size() * 4 / 3);
		for (STableColumn tbCol0 : tbCols) {
			final String tbColName = tbCol0.getName().toUpperCase(); 
			if (tbColName.equalsIgnoreCase(oldName)) {
				tbCol = tbCol0;
			} else if (tbColName.equalsIgnoreCase(newName)) {
				tbColDup = tbCol0;
			} else {
				set.add(tbColName);
			}
		}
		if (tbCol == null) return null;
		
		String newName0 = null;
		if (newName == null) {
			// Generate a newer name if want to clear the cell
			newName0 = "Column";
			final String newNameUpper = newName0.toUpperCase();
			for (int j = tbCols.size(); j > 0; --j) {
				if (!set.contains(newNameUpper + j)) {
					newName0 = newName0 + j;
					break;
				}
			}
		} else if (tbColDup != null) {
			// Generate a newer name if found duplicate new name;
			newName0 = newName;
			final String newNameUpper = newName0.toUpperCase();
			for (int j = 2, len = tbCols.size() + 2; j < len; ++j) {
				if (!set.contains(newNameUpper + j)) {
					newName0 = newName0 + j;
					break;
				}
			}
		}
		
		final String newName1 = newName0 != null ? newName0 : newName; 
		tbCol.setName(newName1);
		
		renameColumnNameFormula(table, oldName, newName1);
		
		return newName0 != null ? newName0 : null;
	}
	
	//ZSS-967
	private void renameColumnNameFormula(STable table, String oldName, String newName) {
		AbstractBookSeriesAdv bs = (AbstractBookSeriesAdv)getBookSeries();
		FormulaTunerHelper tuner = new FormulaTunerHelper(bs);
		DependencyTable dt = bs.getDependencyTable();
		final String tableName = table.getName();
		final Ref ref = new ColumnPrecedentRefImpl(table.getBook().getBookName(), tableName, oldName); //old name
		Set<Ref> dependents = dt.getDirectDependents(ref);
		if(dependents.size()>0){
			//clear the dependents dependency before rename it's Name name
			for(Ref dependent:dependents){
				dt.clearDependents(dependent);
			}
			
			//rebuild the the formula by tuner
			tuner.renameColumnName(table, oldName, newName, dependents);
		}
	}
	
	//ZSS-1041
	@Override
	public SCellStyle getOrCreateDefaultHyperlinkStyle() {
		return getOrCreateDefaultHyperlinkStyle(null);
	}
	//ZSS-1238
	@Override
	public SCellStyle getOrCreateDefaultHyperlinkStyle(SCell cell) {
		final SFont defaultFont = this.getDefaultFont();
		final FontMatcher fontMatcher = new FontMatcher(defaultFont);
		fontMatcher.setColor(ColorImpl.BLUE.getHtmlColor());
		fontMatcher.setUnderline(SFont.Underline.SINGLE);
		SFont linkFont = this.searchFont(fontMatcher);
		
		if (linkFont == null) {
			linkFont = this.createFont(defaultFont, true);
			linkFont.setColor(ColorImpl.BLUE);
			linkFont.setUnderline(SFont.Underline.SINGLE);
		}
		final SCellStyle baseStyle = cell == null ? this.getDefaultCellStyle() : cell.getCellStyle(); //ZSS-1238
		final CellStyleMatcher matcher = new CellStyleMatcher(baseStyle);
		matcher.setFont(linkFont);
		SCellStyle linkStyle = this.searchCellStyle(matcher);
		if (linkStyle == null) {
			linkStyle = StyleUtil.cloneCellStyle(this, baseStyle); //will store into book's styleTable
			linkStyle.setFont(linkFont);
		}
		return linkStyle;
	}
	
	//ZSS-1132
	@Override
	public void initDefaultFont() {
		_defaultFont = (AbstractFontAdv) _defaultCellStyles.get(0).getFont();
	}
	
	//ZSS-1132: default character width depends on default font size
	/**
	 * Office Open XML Part 4: Markup Language Reference 3.3.1.12 col (Column
	 * Width & Formatting) The character width 7 is based on Calibri 11 and
	 * character width 8 is base on Calibri 12.
	 */
	//TODO: The character width is get by experiments 
	@Override
	public int getCharWidth() {
		if (_defaultFont != null) {
			final int pt = _defaultFont.getHeightPoints();
			switch(pt) {
			case 12:
				return 8;
			case 11:
			default:
				return 7;
			}
		}
		return 7;
	}

	//ZSS-1140
	@Override
	public SExtraStyle getExtraStyleAt(int idx) {
		return _extraStyles.get(idx);
	}

	//ZSS-1140
	@Override
	public void addExtraStyle(SExtraStyle extraStyle) {
		_extraStyles.add(extraStyle);
	}

	//ZSS-1140
	@Override
	public List<SExtraStyle> getExtraStyles() {
		return _extraStyles;
	}

	//ZSS-1140
	@Override
	public void clearExtraStyles() {
		_extraStyles.clear();		
	}
	
	//ZSS-1141
	@Override
	public int indexOfExtraStyle(SExtraStyle style) {
		if (_extraStyles == null) return -1;
		int j = 0;
		for (SExtraStyle s : _extraStyles) {
			if (s == style) return j; // 20151103, henrichen: must use "==" instead of equals 
			j++;
		}
		return -1;
	}

	//ZSS-992
	@Override
	public STableStyle getTableStyle(String name) {
		return _tableStyles.get(name);
	}

	//ZSS-992
	@Override
	public void addTableStyle(STableStyle tableStyle) {
		_tableStyles.put(tableStyle.getName(), tableStyle);
	}

	//ZSS-992
	@Override
	public List<STableStyle> getTableStyles() {
		return new ArrayList<STableStyle>(_tableStyles.values());
	}

	//ZSS-992
	@Override
	public void clearTableStyles() {
		_tableStyles.clear();		
	}
	
	//ZSS-992
	@Override
	public void setDefaultPivotStyleName(String name) {
		_defaultPivotStyle = name;
	}
	
	//ZSS-992
	//@since 3.8.3
	@Override
	public String getDefaultPivotStyleName() {
		return _defaultPivotStyle;
	}
	
	//ZSS-992
	@Override
	public void setDefaultTableStyleName(String name) {
		_defaultTableStyle = name;
	}
	
	//ZSS-992
	@Override
	public String getDefaultTableStyleName() {
		return _defaultTableStyle;
	}
	
	//ZSS-1183
	@Override
	/*package*/ SCellStyle getOrCreateCellStyle(SCellStyle src) {
		if(src!=null) {
			Validations.argInstance(src, AbstractCellStyleAdv.class);
		}
		SCellStyle style = null;
		if (src instanceof SExtraStyle) {
			CellStyleMatcher matcher = new CellStyleMatcher(src);
			style = searchExtraStyle(matcher);
			
			if(style==null){
				style = ((AbstractCellStyleAdv)src).createCellStyle(this);
			}
			_extraStyles.add((SExtraStyle)style);
		} else {
			CellStyleMatcher matcher = new CellStyleMatcher(src);
			style = searchCellStyle(matcher);
			
			if(style==null){
				style = ((AbstractCellStyleAdv)src).createCellStyle(this);
			}
			_cellStyles.add(style);
		}
		
		return style;
	}
	
	//ZSS-1183
	@Override
	/*package*/ SFont getOrCreateFont(SFont src) {
		final FontMatcher fontMatcher = new FontMatcher(src);
		SFont dest = this.searchFont(fontMatcher);
		
		if (dest == null) {
			dest = this.createFont(src, true);
			dest.setColor(this.createColor(dest.getColor().getHtmlColor()));
		}
		return dest;
	}
	
	//ZSS-1183
	@Override
	/*package*/ SName getOrCreateName(SName src, String sheetName) {
		final String namename = src.getName();
		SName dest = this.getNameByName(namename, sheetName);
		if (dest == null) {
			dest = this.createName(namename, sheetName);
			dest.setRefersToFormula(src.getRefersToFormula());
		}
		return dest;
	}

	//ZSS-1183
	private STableStyleElem _getOrClone(STableStyleElem src) {
		return (STableStyleElem) (src == null ? null :
				((AbstractCellStyleAdv)src).cloneCellStyle(this)); 
	}
	
	//ZSS-1183
	//@since 3.9.0
	@Override
	/*package*/ STableStyle getOrCreateTableStyle(STableStyle src) {
		final String name = src.getName();
		STableStyle tableStyle = EngineFactory.getInstance()
				.createFormatEngine().getExistTableStyle(this, name);
		if (tableStyle == null) {
			final STableStyleElem wholeTable = _getOrClone(src.getWholeTableStyle());
			final STableStyleElem colStripe1 = _getOrClone(src.getColStripe1Style());
			final STableStyleElem colStripe2 = _getOrClone(src.getColStripe2Style());
			final STableStyleElem rowStripe1 = _getOrClone(src.getRowStripe1Style());
			final STableStyleElem rowStripe2 = _getOrClone(src.getRowStripe2Style());
			final STableStyleElem lastCol = _getOrClone(src.getLastColumnStyle());
			final STableStyleElem firstCol = _getOrClone(src.getFirstColumnStyle());
			final STableStyleElem headerRow = _getOrClone(src.getHeaderRowStyle());
			final STableStyleElem totalRow = _getOrClone(src.getTotalRowStyle());
			final STableStyleElem firstHeaderCell = _getOrClone(src.getFirstHeaderCellStyle());
			final STableStyleElem lastHeaderCell = _getOrClone(src.getLastHeaderCellStyle());
			final STableStyleElem firstTotalCell = _getOrClone(src.getFirstTotalCellStyle());
			final STableStyleElem lastTotalCell = _getOrClone(src.getLastTotalCellStyle()); 
			
			tableStyle = new TableStyleImpl(
					name,
					wholeTable,
					colStripe1,
					src.getColStripe1Size(),
					colStripe2,
					src.getColStripe2Size(),
					rowStripe1,
					src.getRowStripe1Size(),
					rowStripe2,
					src.getRowStripe2Size(),
					lastCol,
					firstCol,
					headerRow,
					totalRow,
					firstHeaderCell,
					lastHeaderCell,
					firstTotalCell,
					lastTotalCell);
			
			addTableStyle(tableStyle);
		}
		
		return tableStyle;
	}
	//ZSS-1283
	public boolean isPostProcessing() {
		return _inPostProcessing;
	}
	//ZSS-1283
	public void setPostProcessing(boolean b) {
		_inPostProcessing = b;
	}
}
