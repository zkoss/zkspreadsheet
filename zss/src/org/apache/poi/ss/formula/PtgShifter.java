/* FormulaShifter.java

	Purpose:
		
	Description:
		
	History:
		Jun 2, 2010 2:31:25 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.apache.poi.ss.formula;

import org.apache.poi.hssf.record.formula.Area2DPtgBase;
import org.apache.poi.hssf.record.formula.Area3DPtg;
import org.apache.poi.hssf.record.formula.AreaErrPtg;
import org.apache.poi.hssf.record.formula.AreaPtg;
import org.apache.poi.hssf.record.formula.AreaPtgBase;
import org.apache.poi.hssf.record.formula.DeletedArea3DPtg;
import org.apache.poi.hssf.record.formula.DeletedRef3DPtg;
import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.record.formula.Ref3DPtg;
import org.apache.poi.hssf.record.formula.RefErrorPtg;
import org.apache.poi.hssf.record.formula.RefPtg;
import org.apache.poi.hssf.record.formula.RefPtgBase;
import org.apache.poi.ss.SpreadsheetVersion;
import org.zkoss.zss.model.impl.BookHelper;

/**
 * @author henrichen
 *
 */
public class PtgShifter {

	/**
	 * Extern sheet index of sheet where moving is occurring
	 */
	private final int _externSheetIndex;
	private final int _firstRow;
	private final int _lastRow;
	private final int _rowAmount;
	private final int _firstCol;
	private final int _lastCol;
	private final int _colAmount;
	private final SpreadsheetVersion _ver;

	public PtgShifter(int externSheetIndex, int firstRow, int lastRow, int rowAmount, int firstCol, int lastCol, int colAmount, SpreadsheetVersion ver) {
		if (firstRow > lastRow) {
			throw new IllegalArgumentException("firstRow, lastRow out of order");
		}
		if (firstCol > lastCol) {
			throw new IllegalArgumentException("firstCol, lastCol out of order");
		}
		_externSheetIndex = externSheetIndex;
		_firstRow = firstRow;
		_lastRow = lastRow;
		_rowAmount = rowAmount;
		_firstCol = firstCol;
		_lastCol = lastCol;
		_colAmount = colAmount;
		_ver = ver;
	}
	
	public String toString() {
		StringBuffer sb = new StringBuffer();

		sb.append(getClass().getName());
		sb.append(" [");
		sb.append(_firstRow);
		sb.append(",").append(_lastRow);
		sb.append(",").append(_rowAmount);
		sb.append(",").append(_firstCol);
		sb.append(",").append(_lastCol);
		sb.append(",").append(_colAmount);
		return sb.toString();
	}

	/**
	 * @param ptgs - if necessary, will get modified by this method
	 * @param currentExternSheetIx - the extern sheet index of the sheet that contains the formula being adjusted
	 * @return <code>true</code> if a change was made to the formula tokens
	 */
	public boolean adjustFormula(Ptg[] ptgs, int currentExternSheetIx) {
		if (_rowAmount == 0 && _colAmount == 0) {
			return false;
		}
		boolean refsWereChanged = false;
		for(int i=0; i<ptgs.length; i++) {
			Ptg newPtg = adjustPtg(ptgs[i], currentExternSheetIx);
			if (newPtg != null) {
				refsWereChanged = true;
				ptgs[i] = newPtg;
			}
		}
		return refsWereChanged;
	}

	private Ptg adjustPtg(Ptg ptg, int currentExternSheetIx) {
		if(ptg instanceof RefPtg) {
			if (currentExternSheetIx != _externSheetIndex) {
				// local refs on other sheets are unaffected
				return null;
			}
			final RefPtg rptg = (RefPtg)ptg;
			final Ptg xptg = _rowAmount != 0 ? rowMoveRefPtg(rptg) : rptg; 
			return _colAmount != 0 && xptg instanceof RefPtg ? colMoveRefPtg((RefPtg)xptg) : xptg;
		}
		if(ptg instanceof Ref3DPtg) {
			final Ref3DPtg rptg = (Ref3DPtg)ptg;
			if (_externSheetIndex != rptg.getExternSheetIndex()) {
				// only move 3D refs that refer to the sheet with cells being moved
				// (currentExternSheetIx is irrelevant)
				return null;
			}
			final Ptg xptg = _rowAmount != 0 ? rowMoveRefPtg(rptg) : rptg;
			return _colAmount != 0 && xptg instanceof Ref3DPtg ? colMoveRefPtg((Ref3DPtg)xptg) : xptg;
		}
		if(ptg instanceof Area2DPtgBase) {
			if (currentExternSheetIx != _externSheetIndex) {
				// local refs on other sheets are unaffected
				return ptg;
			}
			final Ptg xptg = _rowAmount != 0 ? rowMoveAreaPtg((Area2DPtgBase)ptg) : ptg;
			return _colAmount != 0 && xptg instanceof Area2DPtgBase ? colMoveAreaPtg((Area2DPtgBase)xptg) : xptg;
		}
		if(ptg instanceof Area3DPtg) {
			final Area3DPtg aptg = (Area3DPtg)ptg;
			if (_externSheetIndex != aptg.getExternSheetIndex()) {
				// only move 3D refs that refer to the sheet with cells being moved
				// (currentExternSheetIx is irrelevant)
				return null;
			}
			final Ptg xptg = _rowAmount != 0 ? rowMoveAreaPtg(aptg) : aptg;
			return _colAmount != 0 && xptg instanceof Area3DPtg ? colMoveAreaPtg((Area3DPtg)xptg) : xptg;
		}
		return null;
	}

	private Ptg rowMoveRefPtg(RefPtgBase rptg) {
		int refRow = rptg.getRow();
		if (_firstRow <= refRow && refRow <= _lastRow) {
			// Rows being moved completely enclose the ref.
			// - move the area ref along with the rows regardless of destination
			//rptg.setRow(refRow + _amountToMove);
			//return rptg;
			return rptgSetRow(rptg, refRow + _rowAmount);
		}
		// else rules for adjusting area may also depend on the destination of the moved rows

		int destFirstRowIndex = _firstRow + _rowAmount;
		int destLastRowIndex = _lastRow + _rowAmount;

		// ref is outside source rows
		// check for clashes with destination

		if (destLastRowIndex < refRow || refRow < destFirstRowIndex) {
			// destination rows are completely outside ref
			return null;
		}

		if (destFirstRowIndex <= refRow && refRow <= destLastRowIndex) {
			// destination rows enclose the area (possibly exactly)
			return createDeletedRef(rptg);
		}
		throw new IllegalStateException("Situation not covered: (" + _firstRow + ", " +
					_lastRow + ", " + _rowAmount + ", " + refRow + ", " + refRow + ")");
	}
	private Ptg rowMoveAreaPtg(AreaPtgBase aptg) {
		int aFirstRow = aptg.getFirstRow();
		int aLastRow = aptg.getLastRow();
		if (_firstRow <= aFirstRow && aLastRow <= _lastRow) {
			// Rows being moved completely enclose the area ref.
			// - move the area ref along with the rows regardless of destination
			//aptg.setFirstRow(aFirstRow + _amountToMove);
			//aptg.setLastRow(aLastRow + _amountToMove);
			//return aptg;
			aptgSetLastRow(aptg, aLastRow + _rowAmount);
			return aptgSetFirstRow(aptg, aFirstRow + _rowAmount);
		}
		// else rules for adjusting area may also depend on the destination of the moved rows

		int destFirstRowIndex = _firstRow + _rowAmount;
		int destLastRowIndex = _lastRow + _rowAmount;

		if (aFirstRow < _firstRow && _lastRow < aLastRow) {
			// Rows moved were originally *completely* within the area ref

			// If the destination of the rows overlaps either the top
			// or bottom of the area ref there will be a change
			if (destFirstRowIndex < aFirstRow && aFirstRow <= destLastRowIndex) {
				// truncate the top of the area by the moved rows
				//aptg.setFirstRow(destLastRowIndex+1);
				//return aptg;
				return aptgSetFirstRow(aptg, destLastRowIndex+1);
			} else if (destFirstRowIndex <= aLastRow && aLastRow < destLastRowIndex) {
				// truncate the bottom of the area by the moved rows
				//aptg.setLastRow(destFirstRowIndex-1);
				aptgSetLastRow(aptg, destFirstRowIndex-1);
				return aptg;
			}
			// else - rows have moved completely outside the area ref,
			// or still remain completely within the area ref
			return null; // - no change to the area
		}
		if (_firstRow <= aFirstRow && aFirstRow <= _lastRow) {
			// Rows moved include the first row of the area ref, but not the last row
			// btw: (aLastRow > _lastMovedIndex)
			if (_rowAmount < 0) {
				// simple case - expand area by shifting top upward
				//aptg.setFirstRow(aFirstRow + _amountToMove);
				//return aptg;
				return aptgSetFirstRow(aptg, aFirstRow + _rowAmount);
			}
			if (destFirstRowIndex > aLastRow) {
				// in this case, excel ignores the row move
				return null;
			}
			int newFirstRowIx = aFirstRow + _rowAmount;
			if (destLastRowIndex < aLastRow) {
				// end of area is preserved (will remain exact same row)
				// the top area row is moved simply
				//aptg.setFirstRow(newFirstRowIx);
				//return aptg;
				return aptgSetFirstRow(aptg, newFirstRowIx);
			}
			// else - bottom area row has been replaced - both area top and bottom may move now
			int areaRemainingTopRowIx = _lastRow + 1;
			if (destFirstRowIndex > areaRemainingTopRowIx) {
				// old top row of area has moved deep within the area, and exposed a new top row
				newFirstRowIx = areaRemainingTopRowIx;
			}
			//aptg.setFirstRow(newFirstRowIx);
			//aptg.setLastRow(Math.max(aLastRow, destLastRowIndex));
			//return aptg;
			aptgSetLastRow(aptg, Math.max(aLastRow, destLastRowIndex));
			return aptgSetFirstRow(aptg, newFirstRowIx);
		}
		if (_firstRow <= aLastRow && aLastRow <= _lastRow) {
			// Rows moved include the last row of the area ref, but not the first
			// btw: (aFirstRow < _firstMovedIndex)
			if (_rowAmount > 0) {
				// simple case - expand area by shifting bottom downward
				//aptg.setLastRow(aLastRow + _amountToMove);
				aptgSetLastRow(aptg, aLastRow + _rowAmount);
				return aptg;
			}
			if (destLastRowIndex < aFirstRow) {
				// in this case, excel ignores the row move
				return null;
			}
			int newLastRowIx = aLastRow + _rowAmount;
			if (destFirstRowIndex > aFirstRow) {
				// top of area is preserved (will remain exact same row)
				// the bottom area row is moved simply
				//aptg.setLastRow(newLastRowIx);
				aptgSetLastRow(aptg, newLastRowIx);
				return aptg;
			}
			// else - top area row has been replaced - both area top and bottom may move now
			int areaRemainingBottomRowIx = _firstRow - 1;
			if (destLastRowIndex < areaRemainingBottomRowIx) {
				// old bottom row of area has moved up deep within the area, and exposed a new bottom row
				newLastRowIx = areaRemainingBottomRowIx;
			}
			//aptg.setFirstRow(Math.min(aFirstRow, destFirstRowIndex));
			//aptg.setLastRow(newLastRowIx);
			//return aptg;
			aptgSetLastRow(aptg, newLastRowIx);
			return aptgSetFirstRow(aptg, Math.min(aFirstRow, destFirstRowIndex));
		}
		// else source rows include none of the rows of the area ref
		// check for clashes with destination

		if (destLastRowIndex < aFirstRow || aLastRow < destFirstRowIndex) {
			// destination rows are completely outside area ref
			return null;
		}

		if (destFirstRowIndex <= aFirstRow && aLastRow <= destLastRowIndex) {
			// destination rows enclose the area (possibly exactly)
			return createDeletedRef(aptg);
		}

		if (aFirstRow <= destFirstRowIndex && destLastRowIndex <= aLastRow) {
			// destination rows are within area ref (possibly exact on top or bottom, but not both)
			return null; // - no change to area
		}

		if (destFirstRowIndex < aFirstRow && aFirstRow <= destLastRowIndex) {
			// dest rows overlap top of area
			// - truncate the top
			//aptg.setFirstRow(destLastRowIndex+1);
			//return aptg;
			return aptgSetFirstRow(aptg, destLastRowIndex+1);
		}
		if (destFirstRowIndex < aLastRow && aLastRow <= destLastRowIndex) {
			// dest rows overlap bottom of area
			// - truncate the bottom
			//aptg.setLastRow(destFirstRowIndex-1);
			aptgSetLastRow(aptg, destFirstRowIndex-1);
			return aptg;
		}
		throw new IllegalStateException("Situation not covered: (" + _firstRow + ", " +
					_lastRow + ", " + _rowAmount + ", " + aFirstRow + ", " + aLastRow + ")");
	}

	private static Ptg createDeletedRef(Ptg ptg) {
		return BookHelper.createDeletedRef(ptg);
	}
	
	private Ptg rptgSetRow(RefPtgBase rptg, int rowNum) {
		if (rowNum > _ver.getLastRowIndex()) {
			return createDeletedRef(rptg); //out of bound
		} else {
			rptg.setRow(rowNum);
			return rptg;
		}
	}
	private void aptgSetLastRow(AreaPtgBase aptg, int rowNum) {
		if (rowNum > _ver.getLastRowIndex()) {
			aptg.setLastRow(_ver.getLastRowIndex());
		} else {
			aptg.setLastRow(rowNum);
		}
	}
	private Ptg aptgSetFirstRow(AreaPtgBase aptg, int rowNum) {
		if (rowNum > _ver.getLastRowIndex()) {
			return createDeletedRef(aptg); //out of bound
		} else {
			aptg.setFirstRow(rowNum);
			return aptg;
		}
	}
	private Ptg colMoveRefPtg(RefPtgBase rptg) {
		int refCol = rptg.getColumn();
		if (_firstCol <= refCol && refCol <= _lastCol) {
			// Cols being moved completely enclose the ref.
			// - move the area ref along with the cols regardless of destination
			//rptg.setCol(refCol + _amountToMove);
			//return rptg;
			return rptgSetCol(rptg, refCol + _colAmount);
		}
		// else rules for adjusting area may also depend on the destination of the moved cols

		int destFirstColIndex = _firstCol + _colAmount;
		int destLastColIndex = _lastCol + _colAmount;

		// ref is outside source cols
		// check for clashes with destination

		if (destLastColIndex < refCol || refCol < destFirstColIndex) {
			// destination cols are completely outside ref
			return null;
		}

		if (destFirstColIndex <= refCol && refCol <= destLastColIndex) {
			// destination cols enclose the area (possibly exactly)
			return createDeletedRef(rptg);
		}
		throw new IllegalStateException("Situation not covered: (" + _firstCol + ", " +
					_lastCol + ", " + _colAmount + ", " + refCol + ", " + refCol + ")");
	}
	private Ptg colMoveAreaPtg(AreaPtgBase aptg) {
		int aFirstCol = aptg.getFirstColumn();
		int aLastCol = aptg.getLastColumn();
		if (_firstCol <= aFirstCol && aLastCol <= _lastCol) {
			// Cols being moved completely enclose the area ref.
			// - move the area ref along with the cols regardless of destination
			//aptg.setFirstCol(aFirstCol + _amountToMove);
			//aptg.setLastCol(aLastCol + _amountToMove);
			//return aptg;
			aptgSetLastCol(aptg, aLastCol + _colAmount);
			return aptgSetFirstCol(aptg, aFirstCol + _colAmount);
		}
		// else rules for adjusting area may also depend on the destination of the moved cols

		int destFirstColIndex = _firstCol + _colAmount;
		int destLastColIndex = _lastCol + _colAmount;

		if (aFirstCol < _firstCol && _lastCol < aLastCol) {
			// Cols moved were originally *completely* within the area ref

			// If the destination of the cols overlaps either the top
			// or bottom of the area ref there will be a change
			if (destFirstColIndex < aFirstCol && aFirstCol <= destLastColIndex) {
				// truncate the top of the area by the moved cols
				//aptg.setFirstCol(destLastColIndex+1);
				//return aptg;
				return aptgSetFirstCol(aptg, destLastColIndex+1);
			} else if (destFirstColIndex <= aLastCol && aLastCol < destLastColIndex) {
				// truncate the bottom of the area by the moved cols
				//aptg.setLastCol(destFirstColIndex-1);
				aptgSetLastCol(aptg, destFirstColIndex-1);
				return aptg;
			}
			// else - cols have moved completely outside the area ref,
			// or still remain completely within the area ref
			return null; // - no change to the area
		}
		if (_firstCol <= aFirstCol && aFirstCol <= _lastCol) {
			// Cols moved include the first col of the area ref, but not the last col
			// btw: (aLastCol > _lastMovedIndex)
			if (_colAmount < 0) {
				// simple case - expand area by shifting top upward
				//aptg.setFirstCol(aFirstCol + _amountToMove);
				//return aptg;
				return aptgSetFirstCol(aptg, aFirstCol + _colAmount);
			}
			if (destFirstColIndex > aLastCol) {
				// in this case, excel ignores the col move
				return null;
			}
			int newFirstColIx = aFirstCol + _colAmount;
			if (destLastColIndex < aLastCol) {
				// end of area is preserved (will remain exact same col)
				// the top area col is moved simply
				//aptg.setFirstCol(newFirstColIx);
				//return aptg;
				return aptgSetFirstCol(aptg, newFirstColIx);
			}
			// else - bottom area col has been replaced - both area top and bottom may move now
			int areaRemainingTopColIx = _lastCol + 1;
			if (destFirstColIndex > areaRemainingTopColIx) {
				// old top col of area has moved deep within the area, and exposed a new top col
				newFirstColIx = areaRemainingTopColIx;
			}
			//aptg.setFirstCol(newFirstColIx);
			//aptg.setLastCol(Math.max(aLastCol, destLastColIndex));
			//return aptg;
			aptgSetLastCol(aptg, Math.max(aLastCol, destLastColIndex));
			return aptgSetFirstCol(aptg, newFirstColIx);
		}
		if (_firstCol <= aLastCol && aLastCol <= _lastCol) {
			// Cols moved include the last col of the area ref, but not the first
			// btw: (aFirstCol < _firstMovedIndex)
			if (_colAmount > 0) {
				// simple case - expand area by shifting bottom downward
				//aptg.setLastCol(aLastCol + _amountToMove);
				aptgSetLastCol(aptg, aLastCol + _colAmount);
				return aptg;
			}
			if (destLastColIndex < aFirstCol) {
				// in this case, excel ignores the col move
				return null;
			}
			int newLastColIx = aLastCol + _colAmount;
			if (destFirstColIndex > aFirstCol) {
				// top of area is preserved (will remain exact same col)
				// the bottom area col is moved simply
				//aptg.setLastCol(newLastColIx);
				aptgSetLastCol(aptg, newLastColIx);
				return aptg;
			}
			// else - top area col has been replaced - both area top and bottom may move now
			int areaRemainingBottomColIx = _firstCol - 1;
			if (destLastColIndex < areaRemainingBottomColIx) {
				// old bottom col of area has moved up deep within the area, and exposed a new bottom col
				newLastColIx = areaRemainingBottomColIx;
			}
			//aptg.setFirstCol(Math.min(aFirstCol, destFirstColIndex));
			//aptg.setLastCol(newLastColIx);
			//return aptg;
			aptgSetLastCol(aptg, newLastColIx);
			return aptgSetFirstCol(aptg, Math.min(aFirstCol, destFirstColIndex));
		}
		// else source cols include none of the cols of the area ref
		// check for clashes with destination

		if (destLastColIndex < aFirstCol || aLastCol < destFirstColIndex) {
			// destination cols are completely outside area ref
			return null;
		}

		if (destFirstColIndex <= aFirstCol && aLastCol <= destLastColIndex) {
			// destination cols enclose the area (possibly exactly)
			return createDeletedRef(aptg);
		}

		if (aFirstCol <= destFirstColIndex && destLastColIndex <= aLastCol) {
			// destination cols are within area ref (possibly exact on top or bottom, but not both)
			return null; // - no change to area
		}

		if (destFirstColIndex < aFirstCol && aFirstCol <= destLastColIndex) {
			// dest cols overlap top of area
			// - truncate the top
			//aptg.setFirstCol(destLastColIndex+1);
			//return aptg;
			return aptgSetFirstCol(aptg, destLastColIndex+1);
		}
		if (destFirstColIndex < aLastCol && aLastCol <= destLastColIndex) {
			// dest cols overlap bottom of area
			// - truncate the bottom
			//aptg.setLastCol(destFirstColIndex-1);
			aptgSetLastCol(aptg, destFirstColIndex-1);
			return aptg;
		}
		throw new IllegalStateException("Situation not covered: (" + _firstCol + ", " +
					_lastRow + ", " + _colAmount + ", " + aFirstCol + ", " + aLastCol + ")");
	}
	private Ptg rptgSetCol(RefPtgBase rptg, int colNum) {
		if (colNum > _ver.getLastColumnIndex()) {
			return createDeletedRef(rptg); //out of bound
		} else {
			rptg.setColumn(colNum);
			return rptg;
		}
	}
	private void aptgSetLastCol(AreaPtgBase aptg, int colNum) {
		if (colNum > _ver.getLastColumnIndex()) {
			aptg.setLastColumn(_ver.getLastColumnIndex());
		} else {
			aptg.setLastColumn(colNum);
		}
	}
	private Ptg aptgSetFirstCol(AreaPtgBase aptg, int colNum) {
		if (colNum > _ver.getLastColumnIndex()) {
			return createDeletedRef(aptg); //out of bound
		} else {
			aptg.setFirstColumn(colNum);
			return aptg;
		}
	}
}
