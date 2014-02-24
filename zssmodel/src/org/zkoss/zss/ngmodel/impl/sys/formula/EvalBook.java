/* EvalWorkbook2.java

	Purpose:
		
	Description:
		
	History:
		Nov 15, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula;

import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.EvaluationCell;
import org.zkoss.poi.ss.formula.EvaluationName;
import org.zkoss.poi.ss.formula.EvaluationSheet;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaParsingWorkbook;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.functions.FreeRefFunction;
import org.zkoss.poi.ss.formula.ptg.NamePtg;
import org.zkoss.poi.ss.formula.ptg.NameXPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.xssf.model.IndexedUDFFinder;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NName;
import org.zkoss.zss.ngmodel.NSheet;

/**
 * modified from org.zkoss.poi.xssf.usermodel.XSSFEvaluationWorkbook
 * @author Josh Micich, Pao
 */
public final class EvalBook implements EvaluationWorkbook, FormulaParsingWorkbook /* implements parsing book for ugly typecast in POI OperationEvaluationContext.getDynamicReference() */ {

	private NBook nbook;
	private IndexedUDFFinder udfFinder = new IndexedUDFFinder(UDFFinder.DEFAULT);
	private ParsingBook parsingBook; // create new one when parsing new formula

	public EvalBook(NBook book) {
		this.nbook = book;
		createParsingBook(); // just in case
	}

	private void createParsingBook() {
		this.parsingBook = new ParsingBook(nbook);
	}

	public NBook getNBook() {
		return nbook;
	}

	public Ptg[] getFormulaTokens(EvaluationCell cell) {
		String text = cell.getStringCellValue();
		int sheetIndex = getSheetIndex(cell.getSheet());
		return getFormulaTokens(sheetIndex, text);
	}

	public Ptg[] getFormulaTokens(int sheetIndex, String formula) {
		createParsingBook(); // create new parsing book before parsing formula
		return FormulaParser.parse(formula, parsingBook, FormulaType.CELL, sheetIndex);
	}

	public EvaluationName getName(String name, int sheetIndex) {
		// return parsingBook.getName(name, sheetIndex);
		NName nname = null;
		if(sheetIndex < 0) {
			// find defined name from book
			nname = nbook.getNameByName(name);
		} else {
			// find defined name from sheet
			NSheet sheet = nbook.getSheet(sheetIndex);
			if(sheet != null) {
				nname = nbook.getNameByName(name, sheet.getSheetName());
			}
		}
		if(nname != null) {
			int index = nbook.getNames().indexOf(nname);
			return new EvalName(name, index, nname.getRefersToFormula(), sheetIndex);
		} else {
			return null;
		}
	}

	public EvaluationName getName(NamePtg namePtg) {
		// find defined name from book
		String name = parsingBook.getNameText(namePtg);
		if(name != null) {
			NName nname = nbook.getNameByName(name);
			if(nname != null) {
				return new EvalName(name, namePtg.getIndex(), nname.getRefersToFormula(), -1);
			}
		}
		return new EvalName(name, namePtg.getIndex(), null, -1);
	}

	public ExternalName getExternalName(int externSheetIndex, int externNameIndex) {
		throw new RuntimeException("Not implemented yet"); // TODO do we need this?
	}

	public UDFFinder getUDFFinder() {
		return udfFinder;
	}

	public EvaluationSheet getSheet(int sheetIndex) {
		NSheet sheet = nbook.getSheet(sheetIndex);
		return sheet != null ? new EvalSheet(sheet) : null;
	}

	public int getSheetIndex(EvaluationSheet evalSheet) {
		// return sheet index (not external sheet index)
		if(evalSheet instanceof EvalSheet) {
			NSheet sheet = ((EvalSheet)evalSheet).getNSheet();
			return nbook.getSheetIndex(sheet);
		}
		return -1;
	}

	public int getSheetIndex(String sheetName) {
		NSheet sheet = nbook.getSheetByName(sheetName);
		return sheet != null ? nbook.getSheetIndex(sheet) : -1;
	}

	// ** below was added by ZPOI **

	public String getSheetName(int sheetIndex) {
		NSheet sheet = nbook.getSheet(sheetIndex);
		return sheet != null ? sheet.getSheetName() : null;
	}

	public ExternalSheet getExternalSheet(int externSheetIndex) {
		return parsingBook.getExternalSheet(externSheetIndex);
	}

	public String resolveNameXText(NameXPtg n) {
		// check function name by function finder including built-in function
		String name = parsingBook.resolveNameXText(n);
		FreeRefFunction function = udfFinder.findFunction(name);
		return function != null ? name : null;
	}

	/**
	 * @return the sheet index of the sheet with the given external index.
	 */
	public int convertFromExternSheetIndex(int externSheetIndex) {
		ExternalSheet sheet = parsingBook.getAnyExternalSheet(externSheetIndex);
		if(sheet != null) {
			if(sheet.getWorkbookName() == null) { // must be same book
				return getSheetIndex(sheet.getSheetName());
			}
		}
		return -1;
	}

	public int convertLastIndexFromExternSheetIndex(int externSheetIndex) {
		ExternalSheet sheet = parsingBook.getAnyExternalSheet(externSheetIndex);
		if(sheet != null) {
			if(sheet.getWorkbookName() == null) { // must be same book
				return getSheetIndex(sheet.getLastSheetName());
			}
		}
		return -1;
	}

	/**
	 * name to represent named range
	 * @author Pao
	 */
	private class EvalName implements EvaluationName {

		private final String name;
		private final int nameIndex;
		private String refersToFormula;
		private int sheetIndex;

		/**
		 * @param name
		 * @param nameIndex
		 * @param refersToFormula
		 * @param sheetIndex sheet index; if -1, indicates whole book.
		 */
		public EvalName(String name, int nameIndex, String refersToFormula, int sheetIndex) {
			this.name = name;
			this.nameIndex = nameIndex;
			this.refersToFormula = refersToFormula;
			this.sheetIndex = sheetIndex;
		}

		public NamePtg createPtg() {
			return new NamePtg(nameIndex);
		}

		public Ptg[] getNameDefinition() {
			// DON'T clear parsing cache here, because of we still evaluate formula here
			return FormulaParser.parse(refersToFormula, parsingBook, FormulaType.NAMEDRANGE, sheetIndex);
		}

		public String getNameText() {
			return name;
		}

		public boolean hasFormula() {
			// according to spec. 18.2.5 definedName (Defined Name)
			return !isFunctionName() && refersToFormula != null && refersToFormula.length() > 0;
		}

		public boolean isFunctionName() {
			return false;
		}

		public boolean isRange() {
			return hasFormula(); // TODO - is this right?
		}
	}
	
	/* delegate to parsing book for ugly typecast in POI OperationEvaluationContext.getDynamicReference() */

	@Override
	public NameXPtg getNameXPtg(String name) {
		return parsingBook.getNameXPtg(name);
	}

	@Override
	public int getExternalSheetIndex(String sheetName) {
		return parsingBook.getExternalSheetIndex(sheetName);
	}

	@Override
	public int getExternalSheetIndex(String workbookName, String sheetName) {
		return parsingBook.getExternalSheetIndex(workbookName, sheetName);
	}

	@Override
	public SpreadsheetVersion getSpreadsheetVersion() {
		return parsingBook.getSpreadsheetVersion();
	}

	@Override
	public String getBookNameFromExternalLinkIndex(String externalLinkIndex) {
		return parsingBook.getBookNameFromExternalLinkIndex(externalLinkIndex);
	}

	@Override
	public EvaluationName getOrCreateName(String name, int sheetIndex) {
		return parsingBook.getOrCreateName(name, sheetIndex);
	}
}
