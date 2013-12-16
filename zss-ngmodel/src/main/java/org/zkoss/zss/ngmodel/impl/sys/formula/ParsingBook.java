/* ParsingBook.java

	Purpose:
		
	Description:
		
	History:
		Dec 13, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.EvaluationName;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaParsingWorkbook;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.ptg.NamePtg;
import org.zkoss.poi.ss.formula.ptg.NameXPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;

/**
 * A pseudo formula parsing workbook for parsing only.
 * @author Pao
 */
public class ParsingBook implements FormulaParsingWorkbook {

	private List<String> index2name = new ArrayList<String>();

	public ParsingBook(String... names) {
		index2name.addAll(Arrays.asList(names));
	}

	public String getName(int index) {
		return index2name.get(index);
	}

	@Override
	public EvaluationName getName(String name, int sheetIndex) {
		return getOrCreateName(name, sheetIndex);
	}

	@Override
	public NameXPtg getNameXPtg(String name) {
		int index = index2name.size();
		return new NameXPtg(0, index);
	}

	@Override
	public int getExternalSheetIndex(String sheetName) {
		int index = index2name.size();
		index2name.add(sheetName);
		return index;
	}

	@Override
	public int getExternalSheetIndex(String workbookName, String sheetName) {
		throw new RuntimeException("not implemented yet");
	}

	@Override
	public SpreadsheetVersion getSpreadsheetVersion() {
		// TODO zss 3.5
		return SpreadsheetVersion.EXCEL2007;
	}

	@Override
	public String getBookNameFromExternalLinkIndex(String externalLinkIndex) {
		// TODO
		return null;
	}

	@Override
	public EvaluationName getOrCreateName(String name, int sheetIndex) {
		int nameIndex = index2name.size();
		EvaluationName n = new EvalName(name, nameIndex, null, sheetIndex, false);
		index2name.add(name);
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
			return FormulaParser.parse(refersToFormula, ParsingBook.this, FormulaType.NAMEDRANGE, sheetIndex);
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

}
