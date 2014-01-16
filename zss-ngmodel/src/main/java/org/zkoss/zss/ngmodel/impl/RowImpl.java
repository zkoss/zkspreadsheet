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

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class RowImpl extends AbstractRowAdv {
	private static final long serialVersionUID = 1L;

	private AbstractSheetAdv sheet;
	private int index;
	
	private final IndexPool<AbstractCellAdv> cells = new IndexPool<AbstractCellAdv>(){
		private static final long serialVersionUID = 1L;

		@Override
		void resetIndex(int newidx, AbstractCellAdv obj) {
			obj.setIndex(newidx);
		}};

	private AbstractCellStyleAdv cellStyle;
	
	
	private Integer height;
	private boolean hidden = false;
	private boolean customHeight = false;

	public RowImpl(AbstractSheetAdv sheet, int index) {
		this.sheet = sheet;
		this.index = index;
	}

	@Override
	public NSheet getSheet() {
		checkOrphan();
		return sheet;
	}

	@Override
	public int getIndex() {
		checkOrphan();
		return index;
	}

	@Override
	public boolean isNull() {
		return false;
	}

	@Override
	AbstractCellAdv getCell(int columnIdx, boolean proxy) {
		AbstractCellAdv cellObj = cells.get(columnIdx);
		if (cellObj != null) {
			return cellObj;
		}
		checkOrphan();
		return proxy ? new CellProxy(sheet, getIndex(), columnIdx) : null;
	}

	@Override
	AbstractCellAdv getOrCreateCell(int columnIdx) {
		AbstractCellAdv cellObj = cells.get(columnIdx);
		if (cellObj == null) {
			checkOrphan();
			if(columnIdx>=getSheet().getBook().getMaxColumnSize()){
				throw new IllegalStateException("can't create the cell that exceeds max column size "+getSheet().getBook().getMaxColumnSize());
			}
			cellObj = new CellImpl(this, columnIdx);
			cells.put(columnIdx, cellObj);
		}
		return cellObj;
	}

	@Override
	public int getStartCellIndex() {
		return cells.firstKey();
	}

	@Override
	public int getEndCellIndex() {
		return cells.lastKey();
	}

	@Override
	protected void onModelInternalEvent(ModelInternalEvent event) {
		// TODO Auto-generated method stub

	}

	@Override
	public void clearCell(int start, int end) {
		// clear before move relation
		for (AbstractCellAdv cell : cells.subValues(start, end)) {
			cell.destroy();
		}
		cells.clear(start, end);
	}

	@Override
	public void insertCell(int cellIdx, int size) {
		if (size <= 0)
			return;
		
		cells.insert(cellIdx, size);
		
		//destroy the cell that exceeds the max size
		int max = getSheet().getBook().getMaxColumnSize();
		Collection<AbstractCellAdv> exceeds = new ArrayList<AbstractCellAdv>(cells.subValues(max, Integer.MAX_VALUE));
		if(exceeds.size()>0){
			cells.trim(max);
		}
		for(AbstractCellAdv cell:exceeds){
			cell.destroy();
		}
	}

	@Override
	public void deleteCell(int cellIdx, int size) {
		if (size <= 0)
			return;
		
		// clear before move relation
		for (AbstractCellAdv cell : cells.subValues(cellIdx, cellIdx + size - 1)) {
			cell.destroy();
		}

		cells.delete(cellIdx, size);
	}

	@Override
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	public void destroy() {
		checkOrphan();
		for (AbstractCellAdv cell : cells.values()) {
			cell.destroy();
		}
		sheet = null;
	}

	@Override
	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	@Override
	public NCellStyle getCellStyle(boolean local) {
		if (local || cellStyle != null) {
			return cellStyle;
		}
		checkOrphan();
		return sheet.getBook().getDefaultCellStyle();
	}

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, CellStyleImpl.class);
		this.cellStyle = (CellStyleImpl) cellStyle;
	}

	@Override
	public int getHeight() {
		if(height!=null){
			return height.intValue();
		}
		checkOrphan();
		return getSheet().getDefaultRowHeight();
	}

	@Override
	public boolean isHidden() {
		return hidden;
	}

	@Override
	public void setHeight(int height) {
		this.height = Integer.valueOf(height);
	}

	@Override
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}
	
	@Override
	public boolean isCustomHeight() {
		return customHeight;
	}
	@Override
	public void setCustomHeight(boolean custom) {
		customHeight = custom;
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	public Iterator<AbstractCellAdv> getCellIterator() {
		return Collections.unmodifiableCollection(cells.values()).iterator();
	}

	@Override
	void setIndex(int newidx) {
		this.index = newidx;
	}

	@Override
	void moveCellTo(AbstractRowAdv target, int start, int end) {
		if(!(target instanceof RowImpl)){
			throw new IllegalStateException("not RowImpl, is "+target);
		}
		
		Collection<AbstractCellAdv> toMove = cells.clear(start, end);
		
		for(AbstractCellAdv cell:toMove){
			AbstractCellAdv old = ((RowImpl)target).cells.put(cell.getColumnIndex(), cell);
			cell.setRow(target);
			if(old!=null){
				old.destroy();
			}
		}
		
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("Row:").append(getIndex()).append(cells.keySet());
		return sb.toString();
	}
}
