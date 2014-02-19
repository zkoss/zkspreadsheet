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
package org.zkoss.zss.ngapi.impl;

import java.util.*;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.lang.Strings;
import org.zkoss.util.Locales;
import org.zkoss.zss.ngapi.*;
import org.zkoss.zss.ngapi.impl.autofill.AutoFillHelper;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NChart.*;
import org.zkoss.zss.ngmodel.NHyperlink.HyperlinkType;
import org.zkoss.zss.ngmodel.impl.*;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.format.*;
import org.zkoss.zss.ngmodel.sys.input.*;
import org.zkoss.zss.ngmodel.util.*;
/**
 * Only those methods that set cell data, cell style, row (column) style, width, height, and hidden consider 3-D references. 
 * Others don't, just perform on first cell.
 * @author dennis
 * @since 3.5.0
 */
public class NRangeImpl implements NRange {

	private NBook book;
	private final List<EffectedRegion> rangeRefs = new ArrayList<EffectedRegion>(
			1);

	private int _column = Integer.MAX_VALUE;
	private int _row = Integer.MAX_VALUE;
	private int _lastColumn = Integer.MIN_VALUE;
	private int _lastRow = Integer.MIN_VALUE;

	public NRangeImpl(NBook book) {
		this.book = book;
	}
	
	public NRangeImpl(NSheet sheet) {
		addRangeRef(sheet, 0, 0, sheet.getBook().getMaxRowIndex(), sheet
				.getBook().getMaxColumnIndex());
	}

	public NRangeImpl(NSheet sheet, int row, int col) {
		addRangeRef(sheet, row, col, row, col);
	}

	public NRangeImpl(NSheet sheet, int tRow, int lCol, int bRow, int rCol) {
		addRangeRef(sheet, tRow, lCol, bRow, rCol);
	}
	
	private NRangeImpl(Collection<SheetRegion> regions) {
		for(SheetRegion region:regions){
			addRangeRef(region.getSheet(), region.getRow(), region.getColumn(), region.getLastRow(), region.getLastColumn());
		}
	}

	private void addRangeRef(NSheet sheet, int tRow, int lCol, int bRow,
			int rCol) {
		Validations.argNotNull(sheet);
		//TODO to support multiple sheet
		rangeRefs.add(new EffectedRegion(sheet, tRow, lCol, bRow, rCol));

		_column = Math.min(_column, lCol);
		_row = Math.min(_row, tRow);
		_lastColumn = Math.max(_lastColumn, rCol);
		_lastRow = Math.max(_lastRow, bRow);

	}
	
	
	public ReadWriteLock getLock(){
		return getBookSeries().getLock();
	}

	
	private class CellVisitorTask extends ReadWriteTask{
		private CellVisitor visitor;
		private boolean stop = false;
		
		private CellVisitorTask(CellVisitor visitor){
			this.visitor = visitor;
		}

		@Override
		public Object invoke() {
			travelCells(visitor);
			return null;
		}
	}
	
	NBookSeries getBookSeries(){
		return getBook().getBookSeries();
	}
	
	NBook getBook(){
		if(book==null){
			book = getSheet().getBook();
		}
		return book;
	}
	
	@Override
	public NSheet getSheet() {
		if(rangeRefs.size()<=0){
			throw new IllegalStateException("can find any effected range");
		}
		return rangeRefs.get(0).sheet;
	}
	@Override
	public int getRow() {
		return _row;
	}
	@Override
	public int getColumn() {
		return _column;
	}
	@Override
	public int getLastRow() {
		return _lastRow;
	}
	@Override
	public int getLastColumn() {
		return _lastColumn;
	}

	private class EffectedRegion {
		private final NSheet sheet;
		private final CellRegion region;

		public EffectedRegion(NSheet sheet, int row, int column, int lastRow,
				int lastColumn) {
			this.sheet = sheet;
			region = new CellRegion(row, column,lastRow,lastColumn);
		}
	}

	static abstract class CellVisitor {
		/**
		 * @param cell
		 * @return true if continue the visit next cell
		 */
		abstract boolean visit(NCell cell);
	}
	

	/**
	 * travels all the cells in this range
	 * @param visitor
	 */
	private void travelCells(CellVisitor visitor) {

		UpdateCollectorWrap updateWrap = new UpdateCollectorWrap(getBookSeries());
		try{
			for (EffectedRegion r : rangeRefs) {
				CellRegion region = r.region;
				loop1:
				for (int i = region.row; i <= region.lastRow; i++) {
					for (int j = region.column; j <= region.lastColumn; j++) {
						NCell cell = r.sheet.getCell(i, j);
						boolean conti = visitor.visit(cell);
						if(!conti){
							break loop1;
						}
					}
				}
			}
		}finally{
			updateWrap.doFinially();
		}
		updateWrap.doNotify();
	}

	private void handleRefNotifyContentChange(NBookSeries bookSeries,HashSet<Ref> notifySet) {
		// notify changes
		new RefNotifyContentChangeHelper(bookSeries).notifyContentChange(notifySet);
	}

	private boolean euqlas(Object obj1, Object obj2) {
		if (obj1 == obj2) {
			return true;
		}
		if (obj1 != null) {
			return obj1.equals(obj2);
		}
		return false;
	}
	@Override
	public void setValue(final Object value) {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				Object cellval = cell.getValue();
				if (!euqlas(cellval, value)) {
					cell.setValue(value);
				}
				return true;
			}
		}).doInWriteLock(getLock());
	}
	
	@Override
	public void clearContents() {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				if (!cell.isNull()) {
					cell.setHyperlink(null);
					cell.clearValue();
				}
				return true;
			}
		}).doInWriteLock(getLock());
	}

	static class ResultWrap<T> {
		T obj;
		public ResultWrap(){}
		public ResultWrap(T obj){
			this.obj = obj;
		}
		public T get() {
			return obj;
		}
		public void set(T obj) {
			this.obj = obj;
		}
	}
	
	@Override
	public void setEditText(final String editText) {
		final InputEngine ie = EngineFactory.getInstance().createInputEngine();
		final ResultWrap<InputResult> input = new ResultWrap<InputResult>();
		final ResultWrap<HyperlinkType> hyperlinkType = new ResultWrap<HyperlinkType>();
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				InputResult result;
				if((result = input.get())==null){
					result = ie.parseInput(editText == null ? ""
						: editText, cell.getCellStyle().getDataFormat(), new InputParseContext(Locales.getCurrent()));
					input.set(result);
					
					//check if a hyperlink
					if(result.getType() == CellType.STRING){
						hyperlinkType.set(getHyperlinkType((String)result.getValue()));
					}
				}
				
				Object cellval = cell.getValue();
				Object resultVal = result.getValue();
				String format = result.getFormat();
				if (euqlas(cellval, resultVal)) {
					return true;
				}
				
				switch (result.getType()) {
				case BLANK:
					cell.clearValue();
					break;
				case BOOLEAN:
					cell.setBooleanValue((Boolean) resultVal);
					break;
				case FORMULA:
					cell.setFormulaValue((String) resultVal);
					break;
				case NUMBER:
					if(resultVal instanceof Date){
						cell.setDateValue((Date)resultVal);
					}else{
						cell.setNumberValue((Double) resultVal);
					}
					break;
				case STRING:
					cell.setStringValue((String) resultVal);
					if(hyperlinkType.get()!=null){
						NHyperlink link = cell.setupHyperlink();
						link.setType(hyperlinkType.get());
						link.setAddress((String)resultVal);
						link.setLabel((String)resultVal);
					}
					break;
				case ERROR:
				default:
					cell.setValue(resultVal);
				}
				
				String oldFormat = cell.getCellStyle().getDataFormat();
				if(format!=null && NCellStyle.FORMAT_GENERAL.equals(oldFormat)){
					//if there is a suggested format and old format is not general
					StyleUtil.setDataFormat(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex(), format);
				}
				return true;
			}
		}).doInWriteLock(getLock());
	}
	
	private NHyperlink.HyperlinkType getHyperlinkType(String address) {
		if (address != null) {
			final String addr = address.toLowerCase(); // ZSS-288: support more scheme according to POI code, see  org.zkoss.poi.ss.formula.functions.Hyperlink
			if (addr.startsWith("http://") || addr.startsWith("https://")) {
				return NHyperlink.HyperlinkType.URL;
			} else if (addr.startsWith("mailto:")) {
				return NHyperlink.HyperlinkType.EMAIL;
			} // ZSS-288: don't support auto-create hyperlink for DOCUMENT and FILE type
		}
		return null;
	}

	@Override
	public String getEditText() {
		final ResultWrap<String> r = new ResultWrap<String>();
		new CellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
				r.set(fe.getEditText(cell, new FormatContext(Locales.getCurrent())));		
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
	}

	@Override
	public void notifyChange() {
		NBookSeries bookSeries = getBookSeries();
		LinkedHashSet<Ref> notifySet = new LinkedHashSet<Ref>();
		for (EffectedRegion r : rangeRefs) {
			String bookName = r.sheet.getBook().getBookName();
			String sheetName = r.sheet.getSheetName();
			CellRegion region = r.region;
			Ref ref = new RefImpl(bookName, sheetName, region.row, region.column,region.lastRow,region.lastColumn);
			notifySet.add(ref);
		}
		handleRefNotifyContentChange(bookSeries,notifySet);
	}
	
	@Override
	public boolean isWholeSheet(){
		return isWholeRow()&&isWholeColumn();
	}

	@Override
	public boolean isWholeRow() {
		return _column<=0 && _lastColumn>=getBook().getMaxColumnIndex();
	}

	@Override
	public NRange getRows() {
		return new NRangeImpl(getSheet(), _row, 0, _lastRow,getBook().getMaxColumnIndex());
	}

	@Override
	public void setRowHeight(final int heightPx) {
		setRowHeight(heightPx,true);
	}
	public void setRowHeight(final int heightPx, final boolean custom) {
		new ReadWriteTask() {
			@Override
			public Object invoke() {
				setRowHeightInLock(heightPx,null,custom);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	private void setRowHeightInLock(Integer heightPx,Boolean hidden, Boolean custom){
		LinkedHashSet<SheetRegion> notifySet = new LinkedHashSet<SheetRegion>();

		for (EffectedRegion r : rangeRefs) {
			int maxcol = r.sheet.getBook().getMaxColumnIndex();
			CellRegion region = r.region;
			
			for (int i = region.row; i <= region.lastRow; i++) {
				NRow row = r.sheet.getRow(i);
				if(heightPx!=null){
					row.setHeight(heightPx);
				}
				if(hidden!=null){
					row.setHidden(hidden);
				}
				if(custom!=null){
					row.setCustomHeight(custom);
				}
				notifySet.add(new SheetRegion(r.sheet,i,0,i,maxcol));
			}
		}

		new NotifyChangeHelper().notifyRowColumnSizeChange(notifySet);
	}

	@Override
	public boolean isWholeColumn() {
		return _row<=0 && _lastRow>=getBook().getMaxRowIndex();
	}

	@Override
	public NRange getColumns() {
		return new NRangeImpl(getSheet(), 0, _column, getBook().getMaxRowIndex(), _lastColumn);
	}

	@Override
	public void setColumnWidth(final int widthPx) {
		setColumnWidth(widthPx,true);
	}
	public void setColumnWidth(final int widthPx,final boolean custom) {
		new ReadWriteTask() {
			@Override
			public Object invoke() {
				setColumnWidthInLock(widthPx,null,custom);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	private void setColumnWidthInLock(Integer widthPx,Boolean hidden, Boolean custom){
		LinkedHashSet<SheetRegion> notifySet = new LinkedHashSet<SheetRegion>();

		for (EffectedRegion r : rangeRefs) {
			int maxrow = r.sheet.getBook().getMaxRowIndex();
			CellRegion region = r.region;
			
			for (int i = region.column; i <= region.lastColumn; i++) {
				NColumn column = r.sheet.getColumn(i);
				if(widthPx!=null){
					column.setWidth(widthPx);
				}
				if(hidden!=null){
					column.setHidden(hidden);
				}
				if(custom!=null){
					column.setCustomWidth(true);
				}
				notifySet.add(new SheetRegion(r.sheet,0,i,maxrow,i));
			}
		}
		new NotifyChangeHelper().notifyRowColumnSizeChange(notifySet);
	}

	@Override
	public NHyperlink getHyperlink() {
		final ResultWrap<NHyperlink> r = new ResultWrap<NHyperlink>();
		new CellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				r.set(cell.getHyperlink());		
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
	}

	@Override
	public NRange copy(final NRange dstRange, final boolean cut) {
		PasteOption option = new PasteOption();
		option.setCut(cut);
		return pasteSpecial0(dstRange,option);		
	}

	@Override
	public NRange copy(NRange dstRange) {
		return copy(dstRange,false);
	}

	@Override
	public NRange pasteSpecial(NRange dstRange, PasteType pasteType,
			PasteOperation pasteOp, boolean skipBlanks, boolean transpose) {
		PasteOption option = new PasteOption();
		option.setSkipBlank(skipBlanks);
		option.setTranspose(transpose);
		option.setPasteType(toModelPasteType(pasteType));
		option.setPasteOperation(toModelPasteOperation(pasteOp));
		return pasteSpecial0(dstRange,option);
	}
	
	private PasteOption.PasteOperation toModelPasteOperation(
			PasteOperation pasteOp) {
		switch(pasteOp){
		case ADD:
			return PasteOption.PasteOperation.ADD;
		case DIV:
			return PasteOption.PasteOperation.DIV;
		case MUL:
			return PasteOption.PasteOperation.MUL;
		case NONE:
			return PasteOption.PasteOperation.NONE;
		case SUB:
			return PasteOption.PasteOperation.SUB;
		}
		throw new IllegalStateException("unknow operation "+pasteOp);
	}

	private PasteOption.PasteType toModelPasteType(
			PasteType pasteType) {
		switch(pasteType){
		case ALL:
			return PasteOption.PasteType.ALL;
		case ALL_EXCEPT_BORDERS:
			return PasteOption.PasteType.ALL_EXCEPT_BORDERS;
		case COLUMN_WIDTHS:
			return PasteOption.PasteType.COLUMN_WIDTHS;
		case COMMENTS:
			return PasteOption.PasteType.COMMENTS;
		case FORMATS:
			return PasteOption.PasteType.FORMATS;
		case FORMULAS:
			return PasteOption.PasteType.FORMULAS;
		case FORMULAS_AND_NUMBER_FORMATS:
			return PasteOption.PasteType.FORMULAS_AND_NUMBER_FORMATS;
		case VALIDATAION:
			return PasteOption.PasteType.VALIDATAION;
		case VALUES:
			return PasteOption.PasteType.VALUES;
		case VALUES_AND_NUMBER_FORMATS:
			return PasteOption.PasteType.VALUES_AND_NUMBER_FORMATS;
		}
		throw new IllegalStateException("unknow type "+pasteType);
	}

	public NRange pasteSpecial0(final NRange dstRange, final PasteOption option) {
		final ResultWrap<CellRegion> effectedRegion = new ResultWrap<CellRegion>();
		return (NRange)new ModelUpdateTask(){
			@Override
			Object doInvokePhase() {
				CellRegion effected = dstRange.getSheet().pasteCell(new SheetRegion(getSheet(),getRow(),getColumn(),getLastRow(),getLastColumn()), 
						new CellRegion(dstRange.getRow(),dstRange.getColumn(),dstRange.getLastRow(),dstRange.getLastColumn()),
						option);
				effectedRegion.set(effected);
				return new NRangeImpl(getSheet(),effected.getRow(),effected.getColumn(),effected.getLastRow(),effected.getLastColumn());
			}
			@Override
			void doNotifyPhase() {
				if(option.getPasteType()==PasteOption.PasteType.COLUMN_WIDTHS){
					CellRegion effected = effectedRegion.get();
					new NotifyChangeHelper().notifyRowColumnSizeChange(new SheetRegion(dstRange.getSheet(),effected));
				}
			}
			
		}.doInWriteLock(getLock());		
		
	}

	@Override
	public void insert(InsertShift shift, InsertCopyOrigin copyOrigin) {
		//TODO
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void delete(DeleteShift shift) {
		//TODO
		throw new UnsupportedOperationException("not implement yet");
	}

	@Override
	public void merge(final boolean across) {
		new ModelUpdateTask(){
			@Override
			Object doInvokePhase() {
				new MergeHelper(NRangeImpl.this).merge(across);
				return null;
			}
			@Override
			void doNotifyPhase() {}
			
		}.doInWriteLock(getLock());
	}

	@Override
	public void unmerge() {
		new ModelUpdateTask(){
			@Override
			Object doInvokePhase() {
				new MergeHelper(NRangeImpl.this).unmerge(true);
				return null;
			}
			@Override
			void doNotifyPhase() {}
			
		}.doInWriteLock(getLock());
	}

	@Override
	public void setBorders(final ApplyBorderType borderType,final BorderType lineStyle,
			final String color) {
		new ModelUpdateTask(){
			@Override
			Object doInvokePhase() {
				new BorderHelper(NRangeImpl.this).applyBorder(borderType,lineStyle,color);
				return null;
			}
			@Override
			void doNotifyPhase() {}
			
		}.doInWriteLock(getLock());
		
	}

	@Override
	public void move(final int nRow, final int nCol) {
		new ModelUpdateTask(){
			@Override
			Object doInvokePhase() {
				NSheet sheet = getSheet();
				sheet.moveCell(getRow(), getColumn(), getLastRow(), getLastColumn(), nRow, nCol); 
				return null;
			}
			@Override
			void doNotifyPhase() {}
			
		}.doInWriteLock(getLock());
	}

	@Override
	public void setCellStyle(final NCellStyle style) {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				cell.setCellStyle(style);
				return true;
			}
		}).doInWriteLock(getLock());
	}
	
	@Override
	public NCellStyle getCellStyle() {
		final ResultWrap<NCellStyle> r = new ResultWrap<NCellStyle>();
		new CellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				r.set(cell.getCellStyle());		
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
	}

	@Override
	public void fill(final NRange dstRange, final FillType fillType) {
		NSheet sheet = getSheet();
		if(!dstRange.getSheet().equals(sheet)){
			throw new InvalidateModelOpException("the source sheet and destination sheet aren't the same");
		}
		new ModelUpdateTask() {
			@Override
			Object doInvokePhase() {
				autoFillInLock(new CellRegion(getRow(),getColumn(),getLastRow(),getLastColumn()),
						new CellRegion(dstRange.getRow(),dstRange.getColumn(),dstRange.getLastRow(),dstRange.getLastColumn()), fillType);
				return null;
			}
			@Override
			void doNotifyPhase() {}
		}.doInWriteLock(getLock());
	}
	
	private void autoFillInLock(CellRegion src,CellRegion dest, FillType fillType){
		NSheet sheet = getSheet();
		new AutoFillHelper().fill(sheet, src,dest, fillType);
	}

	@Override
	public void fillDown() {
		new ModelUpdateTask() {
			@Override
			Object doInvokePhase() {
				autoFillInLock(new CellRegion(getRow(),getColumn(),getLastRow(),getLastColumn()),
						new CellRegion(getRow(),getColumn(),getLastRow(),getLastColumn()), FillType.COPY);
				return null;
			}
			@Override
			void doNotifyPhase() {}
		}.doInWriteLock(getLock());
	}

	@Override
	public void fillLeft() {
		new ModelUpdateTask() {
			@Override
			Object doInvokePhase() {
				autoFillInLock(new CellRegion(getRow(),getLastColumn(),getLastRow(),getLastColumn()),
						new CellRegion(getRow(),getColumn(),getLastRow(),getLastColumn()), FillType.COPY);
				return null;
			}
			@Override
			void doNotifyPhase() {}
		}.doInWriteLock(getLock());
	}

	@Override
	public void fillRight() {
		new ModelUpdateTask() {
			@Override
			Object doInvokePhase() {
				autoFillInLock(new CellRegion(getRow(),getColumn(),getLastRow(),getColumn()),
						new CellRegion(getRow(),getColumn(),getLastRow(),getLastColumn()), FillType.COPY);
				return null;
			}
			@Override
			void doNotifyPhase() {}
		}.doInWriteLock(getLock());
	}

	@Override
	public void fillUp() {
		new ModelUpdateTask() {
			@Override
			Object doInvokePhase() {
				autoFillInLock(new CellRegion(getLastRow(),getColumn(),getLastRow(),getLastColumn()),
						new CellRegion(getRow(),getColumn(),getLastRow(),getLastColumn()), FillType.COPY);
				return null;
			}
			@Override
			void doNotifyPhase() {}
		}.doInWriteLock(getLock());
	}

	@Override
	public void setHidden(final boolean hidden) {
		new ReadWriteTask() {
			@Override
			public Object invoke() {
				setHiddenInLock(hidden);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	
	private boolean isWholeRow(NBook book,CellRegion region){
		return region.column<=0 && region.lastColumn>=book.getMaxColumnIndex();
	}
	
	private boolean isWholeColumn(NBook book,CellRegion region){
		return region.row<=0 && region.lastRow>=book.getMaxRowIndex();
	}

	protected void setHiddenInLock(boolean hidden) {
		LinkedHashSet<SheetRegion> notifySet = new LinkedHashSet<SheetRegion>();
		for (EffectedRegion r : rangeRefs) {
			NBook book = r.sheet.getBook();
			int maxcol = r.sheet.getBook().getMaxColumnIndex();
			int maxrow = r.sheet.getBook().getMaxRowIndex();
			CellRegion region = r.region;
			
			if(isWholeRow(book,region)){//hidden the row when it is whole row
				for(int i = region.getRow(); i<=region.getLastRow();i++){
					NRow row = r.sheet.getRow(i);
					if(row.isHidden()==hidden)
						continue;
					row.setHidden(hidden);
					notifySet.add(new SheetRegion(r.sheet,i,0,i,maxcol));
				}
			}else if(isWholeColumn(book,region)){
				for(int i = region.getColumn(); i<=region.getLastColumn();i++){
					NColumn col = r.sheet.getColumn(i);
					if(col.isHidden()==hidden)
						continue;
					col.setHidden(hidden);
					notifySet.add(new SheetRegion(r.sheet,0,i,maxrow,i));
				}
			}
		}
		new NotifyChangeHelper().notifyRowColumnSizeChange(notifySet);
	}

	@Override
	public void setDisplayGridlines(final boolean show) {
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				for (EffectedRegion r : rangeRefs) {
					NSheet sheet = r.sheet;
					if(sheet.getViewInfo().isDisplayGridline()!=show){
						sheet.getViewInfo().setDisplayGridline(show);
						new NotifyChangeHelper().notifyDisplayGirdline(sheet,show);
					}
				}
				return null;
			}
		}.doInWriteLock(getLock());
		
	}

	@Override
	public void protectSheet(final String password) {
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				for (EffectedRegion r : rangeRefs) {
					NSheet sheet = r.sheet;
					if(sheet.isProtected() && password==null){
						sheet.setPassword(null);
						new NotifyChangeHelper().notifyProtectSheet(sheet,false);
					}else if(!sheet.isProtected() && password!=null){
						sheet.setPassword(password);
						new NotifyChangeHelper().notifyProtectSheet(sheet,true);
					}
				}
				return null;
			}
		}.doInWriteLock(getLock());
	}

	@Override
	public void setHyperlink(final HyperlinkType linkType,final String address,
			final String display) {
		new CellVisitorTask(new CellVisitor() {
			public boolean visit(NCell cell) {
				NHyperlink link = cell.setupHyperlink();
				link.setType(linkType);
				link.setAddress(address);
				link.setLabel(display);
				
				String text = display;
				while(text.startsWith("=")){
					text = text.substring(1);
				}
				cell.setStringValue(text);
				return true;
			}
		}).doInWriteLock(getLock());
	}

	@Override
	public Object getValue() {
		//Dennis, Should I follow the original implementation in BookHelper:getCellValue(Cell cell) ? it doesn't look appropriately to Range api.
		//I make this api become more easier to get the cell value (if it is formula, then get formula evaluation result
		
		final ResultWrap<Object> r = new ResultWrap<Object>();
		new CellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				Object val = cell.getValue();
				r.set(val);
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
	}

	@Override
	public NRange getOffset(int rowOffset, int colOffset) {
		//follow the original XRange implementation
		if (rowOffset == 0 && colOffset == 0) { //no offset, return this
			return this;
		}
		if (rangeRefs != null && !rangeRefs.isEmpty()) {
			final NBook book = getBook();
			final int maxCol = book.getMaxColumnIndex();
			final int maxRow = book.getMaxRowIndex();
			final LinkedHashSet<SheetRegion> nrefs = new LinkedHashSet<SheetRegion>(rangeRefs.size()); 

			for(EffectedRegion ref : rangeRefs) {
				final int left = ref.region.getColumn() + colOffset;
				final int top = ref.region.getRow() + rowOffset;
				final int right = ref.region.getLastColumn() + colOffset;
				final int bottom = ref.region.getLastRow() + rowOffset;
				
				final NSheet refSheet = ref.sheet;
				final int nleft = colOffset < 0 ? Math.max(0, left) : left;  
				final int ntop = rowOffset < 0 ? Math.max(0, top) : top;
				final int nright = colOffset > 0 ? Math.min(maxCol, right) : right;
				final int nbottom = rowOffset > 0 ? Math.min(maxRow, bottom) : bottom;
				
				if (nleft > nright || ntop > nbottom) { //offset out of range
					continue;
				}
				final SheetRegion refAddr = new SheetRegion(refSheet,ntop, nleft, nbottom, nright);
				if (nrefs.contains(refAddr)) { //same area there, next
					continue;
				}
				nrefs.add(refAddr);
			}
			if (nrefs.isEmpty()) {
				return EMPTY_RANGE;
			} else{
				return new NRangeImpl(nrefs);
			}
		}
		return EMPTY_RANGE;
	}

	@Override
	public boolean isAnyCellProtected() {
		return (Boolean)new ReadWriteTask(){
			@Override
			public Object invoke() {
				for (EffectedRegion r : rangeRefs) {
					NSheet sheet = r.sheet;
					if(sheet.isProtected()){
						CellRegion region = r.region;
						for (int i = region.row; i <= region.lastRow; i++) {
							for (int j = region.column; j <= region.lastColumn; j++) {
								NCellStyle style = r.sheet.getCell(i, j).getCellStyle();
								if(style.isLocked()){
									return true;
								}
							}
						}
					}
				}
				return false;
			}
			
		}.doInReadLock(getLock());
	}

	@Override
	public void deleteSheet() {
		final ResultWrap<NSheet> toDeleteSheet = new ResultWrap<NSheet>();
		final ResultWrap<Integer> toDeleteIndex = new ResultWrap<Integer>();
		//it just handle the first ref
		new ModelUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				int sheetCount;
				if((sheetCount = book.getNumOfSheet())<=1){
					throw new InvalidateModelOpException("can't delete last sheet ");
				}
				
				NSheet toDelete = getSheet();
				
				int index = book.getSheetIndex(toDelete);
//				final int newIndex =  index < (sheetCount - 1) ? index : (index - 1);
				
				toDeleteSheet.set(toDelete);
				toDeleteIndex.set(index);
				
				book.deleteSheet(toDelete);
				return null;
			}

			@Override
			void doNotifyPhase() {
				if(toDeleteSheet.get()!=null){
					new NotifyChangeHelper().notifySheetDelete(getBook(), toDeleteSheet.get(), toDeleteIndex.get());
				}
			}
		}.doInWriteLock(getLock());
	}

	@Override
	public NSheet createSheet(final String name) {
		final ResultWrap<NSheet> resultSheet = new ResultWrap<NSheet>();
		//it just handle the first ref
		return (NSheet)new ModelUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				NSheet sheet;
				if (Strings.isBlank(name)) {
					sheet = book.createSheet(nextSheetName());
				} else {
					sheet = book.createSheet(name);
				}
				resultSheet.set(sheet);
				return sheet;
			}

			@Override
			void doNotifyPhase() {
				if(resultSheet.get()!=null){
					new NotifyChangeHelper().notifySheetCreate(resultSheet.get());
				}
			}
		}.doInWriteLock(getLock());	
	}
	
	private String nextSheetName() {
		NBook book = getBook();
		Integer idx = (Integer)book.getAttribute("zss.nextSheetCount");
		int i = idx==null?1:idx;
		HashSet<String> names = new HashSet<String>();
		for (NSheet sheet : getBook().getSheets()) {
			names.add(sheet.getSheetName());
		}
		String base = "Sheet";
		String name = base + i; 
		while (names.contains(name)) {
			name = base+ (++i);
		}
		book.setAttribute("zss.nextSheetCount", Integer.valueOf(i+1));
		return name;
	}

	@Override
	public void setSheetName(final String newname) {
		//it just handle the first ref
		final ResultWrap<NSheet> resultSheet = new ResultWrap<NSheet>();
		final ResultWrap<String> oldName = new ResultWrap<String>();
		new ModelUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				NSheet sheet = getSheet();
				String old = sheet.getSheetName();
				if(old.equals(newname)){
					return null;
				}
				book.setSheetName(sheet, newname);
				resultSheet.set(sheet);
				oldName.set(old);
				return null;
			}

			@Override
			void doNotifyPhase() {
				if(resultSheet.get()!=null){
					new NotifyChangeHelper().notifySheetNameChange(resultSheet.get(),oldName.get());
				}
			}
		}.doInWriteLock(getLock());	
	}

	@Override
	public void setSheetOrder(final int pos) {
		//it just handle the first ref
		final ResultWrap<NSheet> resultSheet = new ResultWrap<NSheet>();
		final ResultWrap<Integer> oldIdx = new ResultWrap<Integer>();
		new ModelUpdateTask() {			
			@Override
			public Object doInvokePhase() {
				NBook book = getBook();
				NSheet sheet = getSheet();
				
				int old = book.getSheetIndex(sheet);
				if(old==pos){
					return null;
				}
				
				//in our new model, we don't use sheet index, so we don't need to clear anything when move it
				book.moveSheetTo(sheet, pos);
				resultSheet.set(sheet);
				oldIdx.set(old);
				return null;
			}

			@Override
			void doNotifyPhase() {
				if(resultSheet.get()!=null){
					new NotifyChangeHelper().notifySheetReorder(resultSheet.get(),oldIdx.get());
				}
			}
		}.doInWriteLock(getLock());	
	}

	@Override
	public void setFreezePanel(final int numOfRow, final int numOfColumn) {
		//first ref only
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NSheetViewInfo viewInfo = getSheet().getViewInfo();
				viewInfo.setNumOfRowFreeze(numOfRow);
				viewInfo.setNumOfColumnFreeze(numOfColumn);
				notifySheetFreezeChange();
				return null;
			}
		}.doInWriteLock(getLock());	
	}
	
	private void notifySheetFreezeChange(){
		new NotifyChangeHelper().notifySheetFreezeChange(getSheet());
	}

	@Override
	public String getCellFormatText() {
		final ResultWrap<String> r = new ResultWrap<String>();
		new CellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
				r.set(fe.format(cell, new FormatContext(Locales.getCurrent())).getText());		
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
	}
	
	@Override
	public String getCellDataFormat() {
		final ResultWrap<String> r = new ResultWrap<String>();
		new CellVisitorTask(new CellVisitor() {
			@Override
			public boolean visit(NCell cell) {
				FormatEngine fe = EngineFactory.getInstance().createFormatEngine();
				r.set(fe.getFormat(cell, new FormatContext(Locales.getCurrent())));		
				return false;
			}
		}).doInReadLock(getLock());
		return r.get();
	}

	@Override
	public boolean isSheetProtected() {
		//TODO do we really need to use lock in such simple call, it looks overkill.
		return (Boolean)new ReadWriteTask(){
			@Override
			public Object invoke() {
				return getSheet().isProtected();
			}}.doInReadLock(getLock());
	}

	@Override
	public NDataValidation validate(final String editText) {
		final ResultWrap<NDataValidation> retrunVal = new ResultWrap<NDataValidation>();
		new CellVisitorTask(new CellVisitor() {
			boolean visit(NCell cell) {
				NDataValidation validation = getSheet().getDataValidation(cell.getRowIndex(), cell.getColumnIndex());
				if(validation!=null){
					if(!new DataValidationHelper(validation).validate(editText,cell.getCellStyle().getDataFormat())){
						retrunVal.set(validation);
						return false;
					}
				}
				return true;
			}
		}).doInReadLock(getLock());
		return retrunVal.get();
	}

	@Override
	public NRange findAutoFilterRange() {
		return (NRange) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				CellRegion region = new DataRegionHelper(NRangeImpl.this).findAutoFilterDataRegion();
				if(region!=null){
					return NRanges.range(getSheet(),region.getRow(),region.getColumn(),region.getLastRow(),region.getLastColumn());
				}
				return null;
			}
		}.doInReadLock(getLock());
	}
	
	Ref getSheetRef(){
		return new RefImpl((AbstractSheetAdv)getSheet());
	}
	Ref getBookRef(){
		return new RefImpl((AbstractSheetAdv)getBook());
	}
	
	private Set<Ref> toSet(Ref ref){
		Set<Ref> refs = new HashSet(1);
		refs.add(ref);
		return refs;
	}
	
	@Override 
	public NAutoFilter enableAutoFilter(final boolean enable){
		//it just handle the first ref
		return (NAutoFilter) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NSheet sheet = getSheet();
				NAutoFilter filter = sheet.getAutoFilter();
				
				if((filter==null && !enable) || (filter!=null && enable)){
					return filter;
				}
				
				filter = new AutoFilterHelper(NRangeImpl.this).enableAutoFilter(enable);
				notifySheetAutoFilterChange();
				return filter;
			}
		}.doInWriteLock(getLock());
	}

	@Override
	public NAutoFilter enableAutoFilter(final int field, final FilterOp filterOp,
			final Object criteria1, final Object criteria2, final Boolean visibleDropDown) {
		//it just handle the first ref
		return (NAutoFilter) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NAutoFilter filter = new AutoFilterHelper(NRangeImpl.this).enableAutoFilter(field, filterOp, criteria1, criteria2, visibleDropDown);
				notifySheetAutoFilterChange();
				return filter;
			}
		}.doInWriteLock(getLock());
	}
	
	public void resetAutoFilter(){
		//it just handle the first ref
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				new AutoFilterHelper(NRangeImpl.this).resetAutoFilter();
				notifySheetAutoFilterChange();
				return null;
			}
		}.doInWriteLock(getLock());		
	}
	
	public void applyAutoFilter(){
		//it just handle the first ref
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				new AutoFilterHelper(NRangeImpl.this).applyAutoFilter();
				notifySheetAutoFilterChange();
				return null;
			}
		}.doInWriteLock(getLock());		
	}
	
	private void notifySheetAutoFilterChange(){
		new NotifyChangeHelper().notifySheetAutoFilterChange(getSheet());
	}

	@Override
	public void notifyCustomEvent(final String customEventName, final Object data, boolean writelock) {
		//it just handle the first ref
		ReadWriteTask task = new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NotifyChangeHelper notifyHelper =  new NotifyChangeHelper();
				for (EffectedRegion r : rangeRefs) {
					NBook book = r.sheet.getBook();
					notifyHelper.notifyCustomEvent(customEventName,r.sheet,data);
				}
				return null;
			}
		};
		if(writelock){
			task.doInWriteLock(getLock());
		}else{
			task.doInReadLock(getLock());
		}
	}
	
	private abstract class ModelUpdateTask extends ReadWriteTask{
		
		abstract Object doInvokePhase();
		abstract void doNotifyPhase();
		
		@Override
		public Object invoke() {
			UpdateCollectorWrap updateWrap = new UpdateCollectorWrap(getBookSeries());
			Object result = null;
			try{
				result = doInvokePhase();
			}finally{
				updateWrap.doFinially();
			}

			doNotifyPhase();
			updateWrap.doNotify();
			
			return result;
		}
	}
	
	private class UpdateCollectorWrap {
		NBookSeries bookSeries;
		
		LinkedHashSet<Ref> refNotifySet;
		LinkedHashSet<MergeUpdate> mergeNotifySet;
		LinkedHashSet<SheetRegion> cellNotifySet;
		
		RefUpdateCollector refCtx;
		MergeUpdateCollector mergeUpdateCtx;
		CellUpdateCollector cellUpdateCtx;
		
		
		RefUpdateCollector oldRefCtx;
		MergeUpdateCollector oldMergeUpdateCtx;
		CellUpdateCollector oldCellUpdateCtx;
		
		FormulaCacheCleaner oldClearer;
		
		public UpdateCollectorWrap(NBookSeries bookSeries){
			this.bookSeries = bookSeries;
			
			refNotifySet = new LinkedHashSet<Ref>();
			mergeNotifySet = new LinkedHashSet<MergeUpdate>();
			cellNotifySet = new LinkedHashSet<SheetRegion>();			
			
			oldRefCtx = RefUpdateCollector.setCurrent(refCtx = new RefUpdateCollector());
			oldMergeUpdateCtx = MergeUpdateCollector.setCurrent(mergeUpdateCtx = new MergeUpdateCollector());
			oldCellUpdateCtx = CellUpdateCollector.setCurrent(cellUpdateCtx = new CellUpdateCollector());

			oldClearer = FormulaCacheCleaner.setCurrent(new FormulaCacheCleaner(bookSeries));
		}
		
		public void doFinially(){
			
			Set<Ref> refs = refCtx.getRefs();
			Set<SheetRegion> cells = new LinkedHashSet<SheetRegion>();
			
			//remove the duplicate update between cell and refs
			for(SheetRegion region:cellUpdateCtx.getCellUpdates()){
				String bookName = region.getSheet().getBook().getBookName();
				String sheetName = region.getSheet().getSheetName();
				Ref ref = new RefImpl(bookName,sheetName,region.getRow(),region.getColumn());
				if(refs.contains(ref)){
					continue;
				}
				cells.add(region);
			}
			
			refNotifySet.addAll(refs);
			mergeNotifySet.addAll(mergeUpdateCtx.getMergeUpdates());
			cellNotifySet.addAll(cells);
			
			RefUpdateCollector.setCurrent(oldRefCtx);
			MergeUpdateCollector.setCurrent(oldMergeUpdateCtx);
			CellUpdateCollector.setCurrent(oldCellUpdateCtx);
			FormulaCacheCleaner.setCurrent(oldClearer);
		}
		
		public void doNotify(){
			if(refNotifySet.size()>0){
				handleRefNotifyContentChange(bookSeries,refNotifySet);
			}
			if(cellNotifySet.size()>0){
				handleCellNotifyContentChange(cellNotifySet);
			}
			if(mergeNotifySet.size()>0){
				handleMergeNotifyChange(mergeNotifySet);
			}
		}
	}	

	@Override
	public NPicture addPicture(final NViewAnchor anchor, final byte[] image, final NPicture.Format format){
		return (NPicture) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NPicture picture = getSheet().addPicture(format, image, anchor);
				new NotifyChangeHelper().notifySheetPictureAdd(getSheet(), picture.getId());
				return picture;
			}
		}.doInWriteLock(getLock());
	}
	
	public void handleMergeNotifyChange(Set<MergeUpdate> mergeNotifySet) {
		new NotifyChangeHelper().notifyMergeChange(mergeNotifySet);
	}

	public void handleCellNotifyContentChange(Set<SheetRegion> cellNotifySet) {
		new NotifyChangeHelper().notifyCellChange(cellNotifySet);
	}

	@Override
	public void deletePicture(final NPicture picture){
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				String pid = picture.getId();
				getSheet().deletePicture(picture);
				new NotifyChangeHelper().notifySheetPictureDelete(getSheet(), pid);
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	@Override
	public void movePicture(final NPicture picture, final NViewAnchor anchor){
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				picture.setAnchor(anchor);
				new NotifyChangeHelper().notifySheetPictureMove(getSheet(), picture.getId());
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	@Override
	public NChart addChart(final NViewAnchor anchor, final NChartType type,	final NChartGrouping grouping, final NChartLegendPosition pos, final boolean isThreeD) {
		return (NChart) new ReadWriteTask() {			
			@Override
			public Object invoke() {
				NChart chart = getSheet().addChart(type, anchor);
				chart.setThreeD(isThreeD); 
				new ChartDataHelper(NRangeImpl.this).fillChartData(chart);
				chart.setGrouping(grouping);
				chart.setLegendPosition(pos);
				new NotifyChangeHelper().notifySheetChartAdd(getSheet(), chart.getId());
				return chart;
			}
		}.doInWriteLock(getLock());
	}
	
	@Override
	public void deleteChart(final NChart chart){
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				getSheet().deleteChart(chart);
				new NotifyChangeHelper().notifySheetChartDelete(getSheet(), chart.getId());
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	@Override
	public void moveChart(final NChart chart, final NViewAnchor anchor) {
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				chart.setAnchor(anchor);
				new NotifyChangeHelper().notifySheetChartUpdate(getSheet(), chart.getId());
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	@Override
	public void updateChart(final NChart chart) {
		new ReadWriteTask() {			
			@Override
			public Object invoke() {
				new NotifyChangeHelper().notifySheetChartUpdate(getSheet(), chart.getId());
				return null;
			}
		}.doInWriteLock(getLock());
	}
	
	private static final NRange EMPTY_RANGE = new NEmptyRange();
}
