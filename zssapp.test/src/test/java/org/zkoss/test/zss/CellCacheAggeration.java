/* CellCacheAggeration.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 6:37:07 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import java.util.ArrayList;

import com.google.common.base.Objects;
import com.google.inject.Inject;
import com.google.inject.assistedinject.Assisted;

/**
 * @author sam
 *
 */
public class CellCacheAggeration extends ArrayList<CellCache> {
	
	private int tRow;
	private int bRow;
	private int lCol;
	private int rCol;
	
	/*package*/ CellCacheAggeration(int tRow, int lCol, int bRow, int rCol) {
		this.tRow = tRow;
		this.bRow = bRow;
		this.lCol = lCol;
		this.rCol = rCol;
	}
	
	public int getTop() {
		return tRow;
	}
	
	public int getBottom() {
		return bRow;
	}
	
	public int getLeft() {
		return lCol;
	}
	
	public int getRight() {
		return rCol;
	}
	
	public CellCache getCellCache(int row, int col) {
		for (CellCache c : this) {
			if (c.getRow() == row && c.getCol() == col) {
				return c;
			}
		}
		return null;
	}

	public void merge(CellCacheAggeration that, CellCache.Field... fields) {
		for (int i = 0; i < this.size(); i++) {
			CellCache c = get(i);
			CellCache from = that.get(i);
			for (CellCache.Field f : fields) {
				switch (f) {
				case CELL_TYPE:
					c.setCellType(from.getCellType());
					break;
				case MERGE:
					c.setMerge(from.getMerge());
					break;
				case TEXT:
					c.setText(from.getText());
					break;
				case EDIT:
					c.setEdit(from.getEdit());
					break;
				case HORIZONTAL_ALIGN:
					c.setHorizontalAlign(from.getHorizontalAlign());
					break;
				case VERTICAL_ALIGN:
					c.setVerticalAlign(from.getVerticalAlign());
					break;
				case FONT_COLOR:
					c.setFontColor(from.getFontColor());
					break;
				case FILL_COLOR:
					c.setFillColor(from.getFillColor());
					break;
				case BOTTOM_BORDER:
					c.setBottomBorder(from.getBottomBorder());
					break;
				case RIGHT_BORDER:
					c.setRightBorder(from.getRightBorder());
					break;
				}
			}
		}
	}
	
	@Override
	public boolean equals(Object obj) {
		if (obj instanceof CellCacheAggeration) {
			CellCacheAggeration that = (CellCacheAggeration)obj;
			if (this.size() != that.size()) {
				return false;
			}
			
			for (int i = 0; i < this.size(); i++) {
				if (!this.get(i).equals(that.get(i))) {
					return false;
				}
			}
			return true;
			
		}
		return false;
	}

	public boolean equals(CellCacheAggeration that, CellCache.EqualCondition... cods) {
		if (this.size() != that.size()) {
			return false;
		}
		
		for (int i = 0; i < this.size(); i++) {
			if (!this.get(i).equals(that.get(i), cods)) {
				return false;
			}
		}
		return true;
	}
	
	@Override
	public String toString() {
		return Objects.toStringHelper(this)
			.add("tRow", tRow)
			.add("lCol", lCol)
			.add("bRow", bRow)
			.add("rCol", rCol)
			.toString();
	}



	public static interface BuilderFactory {
		public CellCacheAggeration.Builder create(@Assisted("tRow") Integer tRow, @Assisted("lCol") Integer lCol, @Assisted("bRow") Integer bRow, @Assisted("rCol") Integer rCol);
	}
	
	public final static class Builder {
		
		private int tRow;
		private int bRow;
		private int lCol;
		private int rCol;
		
		private final CellCache.Factory cellCacheFactory;
		@Inject
		public Builder(@Assisted("tRow") Integer tRow, @Assisted("lCol") Integer lCol, 
				@Assisted("bRow") Integer bRow, @Assisted("rCol") Integer rCol,
				CellCache.Factory cellCacheFactory) {
			this.tRow = tRow;
			this.bRow = bRow;
			this.lCol = lCol;
			this.rCol = rCol;
			
			this.cellCacheFactory = cellCacheFactory;
		}
		
		public int getTopRow() {
			return tRow;
		}
		
		public int getBottomRow() {
			return bRow;
		}
		
		public int getLeftColumn() {
			return lCol;
		}
		
		public int getRightColumn() {
			return rCol;
		}
		
		public Builder left() {
			return new Builder(tRow, lCol - getColumnSize(), bRow, lCol - 1, cellCacheFactory);
		}
		
		public Builder right() {
			return new Builder(tRow, rCol + 1, bRow, rCol + getColumnSize(), cellCacheFactory);
		}
		
		public Builder up() {
			return new Builder(tRow - getRowSize(), lCol, tRow - 1, rCol, cellCacheFactory);
		}
		
		public Builder down() {
			return new Builder(bRow + 1, lCol, bRow + getRowSize(), rCol, cellCacheFactory);
		}
		
		public Builder offset(int row, int col) {
			return new Builder(row, col, row + getRowSize() - 1, col + getColumnSize() - 1, cellCacheFactory);
		}
		
		public Builder expand(int size) {
			return new Builder(tRow - size, lCol - size, bRow + size, rCol + size, cellCacheFactory);
		}
		
		public Builder expandRight(int size) {
			return new Builder(tRow, lCol, bRow, rCol + size, cellCacheFactory);
		}
		
		public Builder expandDown(int size) {
			return new Builder(tRow, lCol, bRow + size, rCol, cellCacheFactory);
		}
		
		public CellCacheAggeration build() {
//			CellCacheAggeration aggeration = new CellCacheAggeration(tRow, lCol, bRow, rCol);
//			for (int r = tRow; r <= bRow; r++) {
//				for (int c = lCol; c <= rCol; c++) {
//					aggeration.add(cellCacheFactory.create(r, c));
//				}
//			}
//			return aggeration;
			return build("all");
		}
		
		public CellCacheAggeration build(String option) {
			CellCacheAggeration aggeration = new CellCacheAggeration(tRow, lCol, bRow, rCol);
			for (int r = tRow; r <= bRow; r++) {
				for (int c = lCol; c <= rCol; c++) {
					aggeration.add(cellCacheFactory.create(r, c, option));
				}
			}
			return aggeration;
		}
		
		private int getRowSize() {
			return bRow - tRow + 1;
		}
		
		private int getColumnSize() {
			return rCol - lCol + 1;
		}
	}
}