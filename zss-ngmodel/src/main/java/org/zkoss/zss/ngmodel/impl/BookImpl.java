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
package org.zkoss.zss.ngmodel.impl;

import java.awt.Color;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.concurrent.atomic.AtomicInteger;

import org.zkoss.lang.Objects;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.ModelEventListener;
import org.zkoss.zss.ngmodel.ModelEvents;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.NColumnArray;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NName;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationContributor;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationContributorContainer;
import org.zkoss.zss.ngmodel.sys.formula.FormulaClearContext;
import org.zkoss.zss.ngmodel.util.CellStyleMatcher;
import org.zkoss.zss.ngmodel.util.FontMatcher;
import org.zkoss.zss.ngmodel.util.SpreadsheetVersion;
import org.zkoss.zss.ngmodel.util.Strings;
import org.zkoss.zss.ngmodel.util.Validations;

/**
 * @author dennis
 * @since 3.5.0
 */
public class BookImpl extends AbstractBookAdv{
	private static final long serialVersionUID = 1L;

	private final String bookName;
	
	private String shareScope;
	
	private NBookSeries bookSeries;
	
	private final List<AbstractSheetAdv> sheets = new LinkedList<AbstractSheetAdv>();
	private List<AbstractNameAdv> names;
	
	private final List<AbstractCellStyleAdv> cellStyles = new LinkedList<AbstractCellStyleAdv>();
	private final AbstractCellStyleAdv defaultCellStyle;
	private final List<AbstractFontAdv> fonts = new LinkedList<AbstractFontAdv>();
	private final AbstractFontAdv defaultFont;
	private final HashMap<AbstractColorAdv,AbstractColorAdv> colors = new LinkedHashMap<AbstractColorAdv,AbstractColorAdv>();
	
	private final static Random random = new Random(System.currentTimeMillis());
	private final static AtomicInteger bookCount = new AtomicInteger();
	private final String bookId;
	
	private final HashMap<String,AtomicInteger> objIdCounter = new HashMap<String,AtomicInteger>();
	private final int maxRowSize = SpreadsheetVersion.EXCEL2007.getMaxRows();
	private final int maxColumnSize = SpreadsheetVersion.EXCEL2007.getMaxColumns();
	
	private EventListenerAdaptor eventListenerAdaptor;
	
	private transient HashMap<String,Object> attributes;
	
	private EvaluationContributor evalContributor;
	
	public BookImpl(String bookName){
		Validations.argNotNull(bookName);
		this.bookName = bookName;
		bookSeries = new SimpleBookSeriesImpl(this);
		fonts.add(defaultFont = new FontImpl());
		cellStyles.add(defaultCellStyle = new CellStyleImpl(defaultFont));
		colors.put(ColorImpl.WHITE,ColorImpl.WHITE);
		colors.put(ColorImpl.BLACK,ColorImpl.BLACK);
		colors.put(ColorImpl.RED,ColorImpl.RED);
		colors.put(ColorImpl.GREEN,ColorImpl.GREEN);
		colors.put(ColorImpl.BLUE,ColorImpl.BLUE);
		
		bookId = ((char)('a'+random.nextInt(26))) + Long.toString(/*System.currentTimeMillis()+*/bookCount.getAndIncrement(), Character.MAX_RADIX) ;
		
		eventListenerAdaptor = new EventListenerAdaptorImpl();
	}
	
	@Override
	public NBookSeries getBookSeries(){
		return bookSeries;
	}
	
	@Override
	public String getBookName(){
		return bookName;
	}
	
	@Override
	public NSheet getSheet(int i){
		return sheets.get(i);
	}
	
	@Override
	public int getNumOfSheet(){
		return sheets.size();
	}
	
	@Override
	public NSheet getSheetByName(String name){
		for(NSheet sheet:sheets){
			if(sheet.getSheetName().equalsIgnoreCase(name)){
				return sheet;
			}
		}
		return null;
	}
	
	@Override
	public NSheet getSheetById(String id){
		for(NSheet sheet:sheets){
			if(sheet.getId().equals(id)){
				return sheet;
			}
		}
		return null;
	}
	
	protected void checkOwnership(NSheet sheet){
		if(!sheets.contains(sheet)){
			throw new InvalidateModelOpException("doesn't has ownership "+ sheet);
		}
	}
	protected void checkOwnership(NName name){
		if(names==null || !names.contains(name)){
			throw new InvalidateModelOpException("doesn't has ownership "+ name);
		}
	}
	
//	protected String suggestSheetName(String basename){
//		int i = 1;
//		HashSet<String> names = new HashSet<String>();
//		for(NSheet sheet:sheets){
//			names.add(sheet.getSheetName());
//		}
//		String name = basename==null?"Sheet 1":basename;
//		while(names.contains(name)){
//			name = basename + " "+i++;
//		};
//		return name;
//	}
	
	@Override
	void onModelInternalEvent(ModelInternalEvent event){
		//implicitly deliver to sheet
		for(AbstractSheetAdv sheet:sheets){
			sheet.onModelInternalEvent(event);
		}
	}
	
	@Override
	public void sendModelInternalEvent(ModelInternalEvent event){
		//publish event to the series.
		for(NBook book:getBookSeries().getBooks()){
			((AbstractBookAdv)book).onModelInternalEvent(event);
		}
		//TODO some internal event could consider to set it to external(model-event)?
	}
	@Override
	public void sendModelEvent(ModelEvent event){
		eventListenerAdaptor.sendModelEvent(event);
	}
	
	@Override
	public NSheet createSheet(String name) {
		return createSheet(name,null);
	}
	
	@Override
	String nextObjId(String type){
		StringBuilder sb = new StringBuilder(bookId);
		sb.append("_").append(type).append("_");
		AtomicInteger i = objIdCounter.get(type);
		if(i==null){
			objIdCounter.put(type, i = new AtomicInteger(0));
		}
		sb.append(i.getAndIncrement());
		return sb.toString();
	}
	
	@Override
	public NSheet createSheet(String name,NSheet src) {
		checkLegalSheetName(name);
		if(src!=null)
			checkOwnership(src);
		

		AbstractSheetAdv sheet = new SheetImpl(this,nextObjId("sheet"));
		if(src instanceof AbstractSheetAdv){
			((AbstractSheetAdv)src).copyTo(sheet);
		}
		((AbstractSheetAdv)sheet).setSheetName(name);
		sheets.add(sheet);
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		
		sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_SHEET_ADDED, 
				ModelEvents.PARAM_SHEET, sheet));
		return sheet;
	}

	@Override
	public void setSheetName(NSheet sheet, String newname) {
		checkLegalSheetName(newname);
		checkOwnership(sheet);
		
		String oldname = sheet.getSheetName();
		((AbstractSheetAdv)sheet).setSheetName(newname);
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		
		sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_SHEET_RENAMED, 
				ModelEvents.PARAM_SHEET, sheet,
				ModelInternalEvents.PARAM_SHEET_OLD_NAME, oldname));
	}

	private void checkLegalSheetName(String name) {
		if(Strings.isBlank(name)){
			throw new InvalidateModelOpException("sheet name '"+name+"' is not legal");
		}
		if(getSheetByName(name)!=null){
			throw new InvalidateModelOpException("sheet name '"+name+"' is dpulicated");
		}
		//TODO
	}
	
	private void checkLegalNameName(String name,String sheetName) {
		if(Strings.isBlank(name)){
			throw new InvalidateModelOpException("name '"+name+"' is not legal");
		}
		if(getNameByName(name,sheetName)!=null){
			throw new InvalidateModelOpException("name '"+name+"' "+(sheetName==null?"":"'"+sheetName+"'")+" is dpulicated");
		}
		if(sheetName!=null && getSheetByName(sheetName)==null){
			throw new InvalidateModelOpException("no such sheet "+sheetName);
		}
		//TODO
	}

	@Override
	public void deleteSheet(NSheet sheet) {
		checkOwnership(sheet);
		
		((AbstractSheetAdv)sheet).destroy();
		
		int index = sheets.indexOf(sheet);
		sheets.remove(index);
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		
		sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_SHEET_DELETED, 
				ModelEvents.PARAM_SHEET, sheet,
				ModelInternalEvents.PARAM_SHEET_OLD_INDEX, index));
	}

	@Override
	public void moveSheetTo(NSheet sheet, int index) {
		checkOwnership(sheet);
		if(index<0|| index>=sheets.size()){
			throw new InvalidateModelOpException("new position out of bound "+sheets.size() +"<>" +index);
		}
		int oldindex = sheets.indexOf(sheet);
		if(oldindex==index){
			return;
		}
		sheets.remove(oldindex);
		sheets.add(index, (AbstractSheetAdv)sheet);
		
		//create formula cache for any sheet, sheet name, position change
		EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		
		sendModelInternalEvent(createModelInternalEvent(ModelInternalEvents.ON_SHEET_MOVED, 
				ModelEvents.PARAM_SHEET, sheet,
				ModelInternalEvents.PARAM_SHEET_OLD_INDEX, oldindex));
	}

	public void dump(StringBuilder builder) {
		for(AbstractSheetAdv sheet:sheets){
			if(sheet instanceof SheetImpl){
				((SheetImpl)sheet).dump(builder);
			}else{
				builder.append("\n").append(sheet);
			}
		}
	}

	@Override
	public NCellStyle getDefaultCellStyle() {
		return defaultCellStyle;
	}

	@Override
	public NCellStyle createCellStyle(boolean inStyleTable) {
		return createCellStyle(null,inStyleTable);
	}

	@Override
	public NCellStyle createCellStyle(NCellStyle src,boolean inStyleTable) {
		if(src!=null){
			Validations.argInstance(src, AbstractCellStyleAdv.class);
		}
		AbstractCellStyleAdv style = new CellStyleImpl(defaultFont);
		if(src!=null){
			style.copyFrom(src);
		}
		
		if(inStyleTable){
			cellStyles.add(style);
		}
		
		return style;
	}
	
	@Override
	public NCellStyle searchCellStyle(CellStyleMatcher matcher) {
		for(NCellStyle style:cellStyles){
			if(matcher.match(style)){
				return style;
			}
		}
		return null;
	}
	
	
	@Override
	public NFont getDefaultFont() {
		return defaultFont;
	}

	@Override
	public NFont createFont(boolean inFontTable) {
		return createFont(null,inFontTable);
	}

	@Override
	public NFont createFont(NFont src,boolean inFontTable) {
		if(src!=null){
			Validations.argInstance(src, AbstractFontAdv.class);
		}
		AbstractFontAdv font = new FontImpl();
		if(src!=null){
			font.copyFrom(src);
		}
		
		if(inFontTable){
			fonts.add(font);
		}
		
		return font;
	}
	
	@Override
	public NFont searchFont(FontMatcher matcher) {
		for(NFont font:fonts){
			if(matcher.match(font)){
				return font;
			}
		}
		return null;
	}
	
	@Override
	public int getMaxRowSize() {
		return maxRowSize;
	}

	@Override
	public int getMaxColumnSize() {
		return maxColumnSize;
	}

	@Override
	public void optimizeCellStyle() {
		HashMap<String,NCellStyle> stylePool = new LinkedHashMap<String,NCellStyle>();
		cellStyles.clear();
		fonts.clear();
		
		NCellStyle defaultStyle = getDefaultCellStyle();
		NFont defaultFont = getDefaultFont();
		stylePool.put(((AbstractCellStyleAdv)defaultStyle).getStyleKey(), defaultStyle);
		
		for(NSheet sheet:sheets){
			Iterator<NRow> rowIter = sheet.getRowIterator(false); 
			while(rowIter.hasNext()){
				NRow row = rowIter.next();
				
				row.setCellStyle(hitStyle(defaultStyle,row.getCellStyle(),stylePool));
				Iterator<NCell> cellIter = sheet.getCellIterator(row.getIndex());
				while(cellIter.hasNext()){
					NCell cell = cellIter.next();
					cell.setCellStyle(hitStyle(defaultStyle,cell.getCellStyle(),stylePool));
				}
			}
			Iterator<NColumnArray> colIter = sheet.getColumnArrayIterator();
			while(colIter.hasNext()){
				NColumnArray colarr = colIter.next();
				colarr.setCellStyle(hitStyle(defaultStyle,colarr.getCellStyle(),stylePool));
			}
		}
		
		cellStyles.addAll((Collection)stylePool.values());
		String key;
		HashMap<String,NFont> fontPool = new LinkedHashMap<String,NFont>();
		
		fontPool.put(((AbstractFontAdv)defaultFont).getStyleKey(), defaultFont);
		for(NCellStyle style:cellStyles){
			NFont font = style.getFont();
			key = ((AbstractFontAdv)font).getStyleKey();
			if(fontPool.get(key)==null){
				fontPool.put(key, font);
			}
		}
		
		fonts.addAll((Collection)fontPool.values());
		
		colors.clear();//color is immutable, just clear it.
	}
	
	
	@SuppressWarnings("unchecked")
	public List<NCellStyle> getCellStyleTable(){
		return Collections.unmodifiableList((List)cellStyles);
	}
	@SuppressWarnings("unchecked")
	public List<NFont> getFontTable(){
		return Collections.unmodifiableList((List)fonts);
	}
	
	private NCellStyle hitStyle(NCellStyle defaultStyle,NCellStyle currSytle,
			HashMap<String, NCellStyle> stylePool) {
		String key;
		NCellStyle hit;
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
		eventListenerAdaptor.addEventListener(listener);
	}
	@Override
	public void removeEventListener(ModelEventListener listener){
		eventListenerAdaptor.removeEventListener(listener);
	}

	@Override
	public Object getAttribute(String name) {
		return attributes==null?null:attributes.get(name);
	}

	@Override
	public Object setAttribute(String name, Object value) {
		if(attributes==null){
			attributes = new HashMap<String, Object>();
		}
		return attributes.put(name, value);
	}

	@Override
	public Map<String, Object> getAttributes() {
		return attributes==null?Collections.EMPTY_MAP:Collections.unmodifiableMap(attributes);
	}

	@Override
	public NColor createColor(byte r, byte g, byte b) {
		AbstractColorAdv newcolor = new ColorImpl(r,g,b);
		AbstractColorAdv color = colors.get(newcolor);//reuse the existed color object
		if(color==null){
			colors.put(newcolor, color = newcolor);
		}
		return color;
	}

	@Override
	public NColor createColor(String htmlColor) {
		AbstractColorAdv newcolor = new ColorImpl(htmlColor);
		AbstractColorAdv color = colors.get(newcolor);//reuse the existed color object
		if(color==null){
			colors.put(newcolor, color = newcolor);
		}
		return color;
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public List<NSheet> getSheets() {
		return Collections.unmodifiableList((List)sheets);
	}
	@Override
	public NName createName(String namename) {
		return createName(namename,null);
	}
	@Override
	public NName createName(String namename,String sheetName) {
		checkLegalNameName(namename,sheetName);

		AbstractNameAdv name = new NameImpl(this,nextObjId("name"));
		name.setName(namename);
		name.setApplyToSheetName(sheetName);
		
		if(names==null){
			names = new LinkedList<AbstractNameAdv>();
		}
		
		names.add(name);
		
//		sendEvent(ModelEvents.ON_NAME_ADDED, 
//				ModelEvents.PARAM_SHEET, sheet);
		return name;
	}

	@Override
	public void setNameName(NName name, String newname) {
		setNameName(name,newname,null);
	}
	public void setNameName(NName name, String newname, String sheetName) {
		checkLegalNameName(newname,sheetName);
		checkOwnership(name);
		
		String oldname = name.getRefersToSheetName();
		((AbstractNameAdv)name).setName(newname);
		((AbstractNameAdv)name).setApplyToSheetName(sheetName);
		
//		sendEvent(ModelEvents.ON_NAME_RENAMED, 
//				ModelEvents.PARAM_SHEET, sheet,
//				ModelEvents.PARAM_SHEET_OLD_NAME, oldname);
	}

	@Override
	public void deleteName(NName name) {
		checkOwnership(name);
		
		((AbstractNameAdv)name).destroy();
		
		int index = names.indexOf(name);
		names.remove(index);
		
//		sendEvent(ModelEvents.ON_NAME_DELETED, 
//				ModelEvents.PARAM_NAME, sheet,
//				ModelEvents.PARAM_SHEET_OLD_INDEX, index);
	}

	@Override
	public int getNumOfName() {
		return names==null?0:names.size();
	}

	@Override
	public NName getName(int idx) {
		if(names==null){
			throw new ArrayIndexOutOfBoundsException(idx);
		}
		return names.get(idx);
	}

	@Override
	public NName getNameByName(String namename) {
		return getNameByName(namename,null);
	}
	public NName getNameByName(String namename,String sheetName) {
		if(names==null)
			return null;
		for(NName name:names){
			if((sheetName==null || sheetName.equals(name.getApplyToSheetName())) 
					&& name.getName().equalsIgnoreCase(namename)){
				return name;
			}
		}
		return null;
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	@Override
	public List<NName> getNames() {
		return names==null?Collections.EMPTY_LIST:Collections.unmodifiableList((List)names);
	}

	@Override
	public int getSheetIndex(NSheet sheet) {
		return sheets.indexOf(sheet);
	}

	@Override
	public void setShareScope(String scope) {
		if(!Objects.equals(this.shareScope,scope)){
			
			if("disable".equals(scope)){
				eventListenerAdaptor.clear();
				return;
			}
			
			if(eventListenerAdaptor.size()>0){
				throw new IllegalStateException("can't change share scope after registed any listener");
			}
			
			this.shareScope = scope;
			eventListenerAdaptor.clear();
			
			if(scope!=null){
				eventListenerAdaptor = new EventQueueListenerAdaptorImpl(scope, bookId);
			}else{
				eventListenerAdaptor = new EventListenerAdaptorImpl();
			}
		}
	}

	@Override
	public String getShareScope() {
		return shareScope;
	}

	@Override
	void setBookSeries(NBookSeries bookSeries) {
		this.bookSeries = bookSeries;
	}

	@Override
	public EvaluationContributor getEvaluationContributor() {
		return evalContributor;
	}

	@Override
	public void setEvaluationContributor(EvaluationContributor contributor) {
		this.evalContributor = contributor;
	}

}
