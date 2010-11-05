/* RefBookDepencyTracker.java

	Purpose:
		
	Description:
		
	History:
		Mar 23, 2010 4:49:57 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.poi.ss.formula;

import org.zkoss.poi.hssf.record.formula.eval.ErrorEval;
import org.zkoss.poi.hssf.record.formula.eval.NameEval;
import org.zkoss.poi.hssf.record.formula.eval.StringEval;
import org.zkoss.poi.hssf.record.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.eval.NotImplementedException;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.xel.XelContext;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.RefSheet;
import org.zkoss.zss.engine.impl.CellRefImpl;
import org.zkoss.zss.engine.impl.DependencyTrackerHelper;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.XelContextHolder;

/**
 * Implementation of formula dependency tracking.
 * @author henrichen
 *
 */
public class DefaultDependencyTracker implements DependencyTracker {
	private final Book _book;
	public DefaultDependencyTracker(Book book) {
		_book = book;
	}
	
	@Override
	public ValueEval addDependency(OperationEvaluationContext ec, ValueEval opResult, boolean eval) {
		if (opResult instanceof LazyAreaEval) {
			final CellRefImpl srcRef = prepareSrcRef(ec);
			if (srcRef != null) {
				final LazyAreaEval ae = (LazyAreaEval) opResult;
				final String refBookName = ae.getBookName();
				final String refSheetName = ae.getSheetName();
				final String refLastSheetName = ae.getLastSheetName();
				final int tRow = ae.getFirstRow();
				final int lCol = ae.getFirstColumn();
				final int bRow = ae.getLastRow();
				final int rCol = ae.getLastColumn();
				
				myAddDependency(srcRef, refBookName, refSheetName, refLastSheetName, tRow, lCol, bRow, rCol, eval);
			}
		} else if (opResult instanceof LazyRefEval) {
			final CellRefImpl srcRef = prepareSrcRef(ec);
			if (srcRef != null) {
				final LazyRefEval ae = (LazyRefEval) opResult;
				final String refBookName = ae.getBookName();
				final String refSheetName = ae.getSheetName();
				final String refLastSheetName = ae.getLastSheetName();
				final int tRow = ae.getRow();
				final int lCol = ae.getColumn();
	
				myAddDependency(srcRef, refBookName, refSheetName, refLastSheetName, tRow, lCol, tRow, lCol, eval);
			}
		} else if (opResult instanceof NameEval) {
			final String fname = ((NameEval)opResult).getFunctionName(); 
			opResult = fname != null && fname.startsWith("_xlfn.") ? opResult : ErrorEval.NAME_INVALID; 
		}
		return opResult;
	}
	
	protected CellRefImpl prepareSrcRef(OperationEvaluationContext ec) {
		final XelContext ctx = XelContextHolder.getXelContext();
		CellRefImpl srcRef = null;
		boolean isOld = false;
		final String srcSheetName = ec.getSheetName();
		final int srcRow = ec.getRowIndex();
		final int srcCol = ec.getColumnIndex();
		final RefBook srcRefBook = BookHelper.getOrCreateRefBook(_book);
		final RefSheet srcRefSheet = srcRefBook.getOrCreateRefSheet(srcSheetName);
		if (ctx != null) {
			final String srcRefKey = srcRefBook.getBookName()+"]"+srcSheetName+"!"+new CellReference(srcRow, srcCol).formatAsString();
			final Object[] refs = (Object[]) ctx.getAttribute(srcRefKey);
			if (refs != null) {
				srcRef = (CellRefImpl) refs[0];
				isOld = ((Boolean)refs[1]).booleanValue();
			} else {
				srcRef = (CellRefImpl) srcRefSheet.getRef(srcRow, srcCol, srcRow, srcCol);
				if (srcRef == null) { // a new evaluated one
					srcRef = (CellRefImpl) srcRefSheet.getOrCreateRef(srcRow, srcCol, srcRow, srcCol);
				} else {
					isOld = !srcRef.getPrecedents().isEmpty();
				}
				ctx.setAttribute(srcRefKey, new Object[] {srcRef, Boolean.valueOf(isOld)});
			}
			if (isOld) { //an old src ref, no need to add dependency
				return null;
			}
		} else {
			srcRef = (CellRefImpl) srcRefSheet.getOrCreateRef(srcRow, srcCol, srcRow, srcCol); 
		}
		return srcRef;
	}
		
	private void myAddDependency(CellRefImpl srcRef, String refBookname, 
		String refSheetname, String refLastSheetName, int tRow, int lCol, int bRow, int rCol, boolean eval) {
		final Book targetBook = BookHelper.getBook(_book, refBookname);
		if (targetBook == null) {
			throw new UiException("cannot find the named book, have you add it in the Books:"+ refBookname);
		}
		final RefBook targetRefBook = BookHelper.getOrCreateRefBook(targetBook);
		
		final int s1 = targetBook.getSheetIndex(refSheetname);
		final int s2 = targetBook.getSheetIndex(refLastSheetName);
		final int sheetIndex1 = Math.min(s1, s2);
		final int sheetIndex2 = Math.max(s1, s2);
		for(int j = sheetIndex1; j <= sheetIndex2; ++j) {
			final String sheetname = targetBook.getSheetName(j);
			final RefSheet targetRefSheet = targetRefBook.getOrCreateRefSheet(sheetname);
			DependencyTrackerHelper.addDependency(srcRef, targetRefSheet, tRow, lCol, bRow, rCol);
		}
	}
}
