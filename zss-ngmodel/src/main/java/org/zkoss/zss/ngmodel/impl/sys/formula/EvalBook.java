/* EvalWorkbook2.java

	Purpose:
		
	Description:
		
	History:
		Nov 15, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.EvaluationCell;
import org.zkoss.poi.ss.formula.EvaluationName;
import org.zkoss.poi.ss.formula.EvaluationSheet;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaParsingWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderingWorkbook;
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
public final class EvalBook implements FormulaRenderingWorkbook, EvaluationWorkbook, FormulaParsingWorkbook {

	private NBook nbook;
	private IndexedUDFFinder udfFinder = new IndexedUDFFinder(UDFFinder.DEFAULT);
	private List<EvaluationSheet> sheets = new ArrayList<EvaluationSheet>();
	private Map<EvaluationSheet, Integer> sheet2index = new HashMap<EvaluationSheet, Integer>();
	private List<String> names = new ArrayList<String>();
	private Map<String, Integer> name2index = new HashMap<String, Integer>();
	private List<EvaluationName> namedRanges = new ArrayList<EvaluationName>();
	private Map<String, EvaluationName> name2namedRanges = new HashMap<String, EvaluationName>();

	public EvalBook(NBook book) {
		this.nbook = book;
		for(NSheet sheet : book.getSheets()) {
			addSheet(sheet.getSheetName(), new EvalSheet(sheet));
		}
		for(NName n : book.getNames()) {
			addNameRange(-1, n.getName(), n.getRefersToFormula());
		}
	}
	
	public NBook getNBook() {
		return nbook;
	}

	public void addSheet(String name, EvaluationSheet sheet) {
		Integer index = this.sheets.size();
		this.sheets.add(sheet); // index to sheet;
		this.sheet2index.put(sheet, index); // sheet to index;
		this.names.add(name); // index to name;
		this.name2index.put(name, index); // name to index;
	}

	private int convertFromExternalSheetIndex(int externSheetIndex) {
		return externSheetIndex;
	}

	/**
	 * @return the sheet index of the sheet with the given external index.
	 */
	public int convertFromExternSheetIndex(int externSheetIndex) {
		// FIXME check these code
		if(externSheetIndex < 0) {
			return externSheetIndex;
		}

		String name = names.get(externSheetIndex);
		int p = name.indexOf(':');
		if(p < 0) {
			return externSheetIndex;
		} else {
			// 3d reference
			String s1 = name.substring(0, p);
			return getSheetIndex(s1);
		}
	}

	/**
	 * @return the external sheet index of the sheet with the given internal
	 *         index. Used by some of the more obscure formula and named range things.
	 *         Fairly easy on XSSF (we think...) since the internal and external
	 *         indices are the same
	 */
	private int convertToExternalSheetIndex(int sheetIndex) {
		return sheetIndex;
	}

	public int getExternalSheetIndex(String sheetName) {
		Integer sheetIndex = name2index.get(sheetName);
		if(sheetIndex != null) {
			return convertToExternalSheetIndex(sheetIndex);
		} else {
			// not existed and is a 3d reference
			int p = sheetName.indexOf(':');
			if(p >= 0) {
				// add new index for it
				int index = names.size();
				names.add(sheetName);
				name2index.put(sheetName, index);
				return index;
			}
		}
		return -1;
	}

	public EvaluationName getName(String name, int sheetIndex) {
		// book should know all name of named ranges
		String key = sheetIndex < 0 ? name : sheetIndex + name;
		EvaluationName n = name2namedRanges.get(key);
		return (n != null) ? n : name2namedRanges.get(name); // search on book if not existed in sheet
	}

	public int getSheetIndex(EvaluationSheet evalSheet) {
		return sheet2index.get(evalSheet);
	}

	public String getSheetName(int sheetIndex) {
		return names.get(sheetIndex);
	}

	public ExternalName getExternalName(int externSheetIndex, int externNameIndex) {
		throw new RuntimeException("Not implemented yet");
	}

	public NameXPtg getNameXPtg(String name) {
		FreeRefFunction func = udfFinder.findFunction(name);
		if(func == null) {
			return null;
		} else {
			return new NameXPtg(0, udfFinder.getFunctionIndex(name));
		}
	}

	public String resolveNameXText(NameXPtg n) {
		int idx = n.getNameIndex();
		return udfFinder.getFunctionName(idx);
	}

	public EvaluationSheet getSheet(int sheetIndex) {
		return sheets.get(sheetIndex);
	}

	public ExternalSheet getExternalSheet(int externSheetIndex) {
		return null;
	}

	public int getExternalSheetIndex(String workbookName, String sheetName) {
		throw new RuntimeException("not implemented yet");
	}

	public int getSheetIndex(String sheetName) {
		return name2index.get(sheetName);
	}

	public String getSheetNameByExternSheet(int externSheetIndex) {
		int sheetIndex = convertFromExternalSheetIndex(externSheetIndex);
		return names.get(sheetIndex);
	}

	public String getNameText(NamePtg namePtg) {
		EvaluationName name = getName(namePtg);
		return (name != null) ? name.getNameText() : null;
	}

	public EvaluationName getName(NamePtg namePtg) {
		// named ranges
		int index = namePtg.getIndex();
		return namedRanges.get(index);
	}

	public Ptg[] getFormulaTokens(EvaluationCell cell) {
		String text = cell.getStringCellValue();
		int sheetIndex = getSheetIndex(cell.getSheet());
		return getFormulaTokens(sheetIndex, text);
	}

	public UDFFinder getUDFFinder() {
		return udfFinder;
	}

	/**
	 * normal name range
	 * @param sheetIndex less than 0 indicates available in whole book.
	 */
	public EvaluationName addNameRange(int sheetIndex, String name, String refersToFormula) {
		return addNameRange(sheetIndex, name, refersToFormula, false);
	}

	/**
	 * for function type name range
	 */
	public EvaluationName addNameRange(int sheetIndex, String name) {
		return addNameRange(sheetIndex, name, null, true);
	}

	private EvaluationName addNameRange(int sheetIndex, String name, String refersToFormula, boolean isFunctionName) {
		int nameIndex = namedRanges.size();
		EvaluationName n = new EvalName(name, nameIndex, refersToFormula, sheetIndex, isFunctionName);
		// n = (EvaluationName)Proxy.newProxyInstance(ClassLoader.getSystemClassLoader(),
		// new Class<?>[]{EvaluationName.class}, new LogHandler(n));
		namedRanges.add(n);
		String key = sheetIndex < 0 ? name : sheetIndex + name;
		name2namedRanges.put(key, n);
		return n;
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
		private boolean isFunctionName;

		/**
		 * @param name
		 * @param nameIndex
		 * @param refersToFormula
		 * @param sheetIndex sheet index; if -1, indicates whole book.
		 * @param isFunctionName true indicates a reference name of user-defined function
		 */
		public EvalName(String name, int nameIndex, String refersToFormula, int sheetIndex, boolean isFunctionName) {
			this.name = name;
			this.nameIndex = nameIndex;
			this.refersToFormula = refersToFormula;
			this.sheetIndex = sheetIndex;
			this.isFunctionName = isFunctionName;
		}

		public Ptg[] getNameDefinition() {
			return FormulaParser.parse(refersToFormula, EvalBook.this, FormulaType.NAMEDRANGE, sheetIndex);
		}

		public String getNameText() {
			return name;
		}

		public boolean hasFormula() {
			// according to spec. 18.2.5 definedName (Defined Name)
			return !isFunctionName() && refersToFormula != null && refersToFormula.length() > 0;
		}

		public boolean isFunctionName() {
			return isFunctionName;
		}

		public boolean isRange() {
			return hasFormula(); // TODO - is this right?
		}

		public NamePtg createPtg() {
			return new NamePtg(nameIndex);
		}
	}

	public SpreadsheetVersion getSpreadsheetVersion() {
		return SpreadsheetVersion.EXCEL2007;
	}

	// ** below was added by ZPOI **

	@Override
	public Ptg[] getFormulaTokens(int sheetIndex, String formula) {

		// FormulaParsingWorkbook fpw = (FormulaParsingWorkbook)Proxy.newProxyInstance(
		// ClassLoader.getSystemClassLoader(), new Class<?>[]{FormulaParsingWorkbook.class},
		// new LogHandler(this));
		//
		// return FormulaParser.parse(formula, fpw, FormulaType.CELL, sheetIndex);
		return FormulaParser.parse(formula, this, FormulaType.CELL, sheetIndex);
	}

	@Override
	public EvaluationName getOrCreateName(String name, int sheetIndex) {
		EvaluationName n = getName(name, sheetIndex);
		if(n != null) {
			return n;
		}
		return addNameRange(sheetIndex, name, null);
	}

	@Override
	public int convertLastIndexFromExternSheetIndex(int externSheetIndex) {
		// FIXME check these code
		if(externSheetIndex < 0) {
			return externSheetIndex;
		}

		String name = names.get(externSheetIndex);
		int p = name.indexOf(':');
		if(p < 0) {
			return externSheetIndex;
		} else {
			// 3d reference
			String s1 = name.substring(p + 1);
			return getSheetIndex(s1);
		}
	}

	@Override
	public String getBookNameFromExternalLinkIndex(String externalLinkIndex) {
		// TODO
		return null;
	}

	@Override
	public String getExternalLinkIndexFromBookName(String bookname) {
		// TODO
		return null;
	}
}
