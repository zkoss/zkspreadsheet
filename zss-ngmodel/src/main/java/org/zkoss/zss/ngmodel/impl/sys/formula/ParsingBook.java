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
import org.zkoss.poi.ss.formula.EvaluationWorkbook.ExternalSheet;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaParsingWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderingWorkbook;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.ptg.NamePtg;
import org.zkoss.poi.ss.formula.ptg.NameXPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;

/**
 * A pseudo formula parsing workbook for parsing only.
 * @author Pao
 */
public class ParsingBook implements FormulaParsingWorkbook, FormulaRenderingWorkbook {

	private List<String> index2name = new ArrayList<String>();
	private NBook book;
	private List<ExternalSheet> index2sheet = new ArrayList<ExternalSheet>();

	public ParsingBook(NBook book, String... names) {
		this.book = book;
		index2name.addAll(Arrays.asList(names));
	}

	@Override
	public EvaluationName getName(String name, int sheetIndex) {
		return getOrCreateName(name, sheetIndex);
	}

	@Override
	public NameXPtg getNameXPtg(String name) {
		// formula function name
		int index = index2name.size();
		index2name.add(name);
		return new NameXPtg(0, index);
	}

	@Override
	public int getExternalSheetIndex(String sheetName) {
		return getExternalSheetIndex(null, sheetName);
	}

	@Override
	public int getExternalSheetIndex(String workbookName, String sheetName) {
		// create new index and check sheet name is 3D or not
		int index = index2sheet.size();
		int p = sheetName.indexOf(':');
		String name = p < 0 ? sheetName : sheetName.substring(0, p);
		String lastName = p < 0 ? sheetName : sheetName.substring(p+1);
		index2sheet.add(new ExternalSheet(workbookName, name, lastName));
		return index;
	}

	@Override
	public SpreadsheetVersion getSpreadsheetVersion() {
		// TODO zss 3.5
		return SpreadsheetVersion.EXCEL2007;
	}

	@Override
	public String getBookNameFromExternalLinkIndex(String externalLinkIndex) {

		try {
			// if external link index is really a index, convert it and find name from records
			int index = Integer.parseInt(externalLinkIndex);
			String[] names = (String[])book.getAttribute(FormulaEngine.KEY_EXTERNAL_BOOK_NAMES);
			if(names != null) {
				return names[index];
			}
		} catch(NumberFormatException e) {
			// do nothing
		} catch(IndexOutOfBoundsException e) {
			// do nothing
		}

		// otherwise, it should be a book name already and just return itself.
		return externalLinkIndex;
	}

	@Override
	public EvaluationName getOrCreateName(String name, int sheetIndex) {
		int nameIndex = index2name.size();
		EvaluationName n = new SimpleName(name, nameIndex, sheetIndex);
		index2name.add(name);
		return n;
	}
	
	/* FormulaRenderingWorkbook */

	@Override
	public String getNameText(NamePtg namePtg) {
		return index2name.get(namePtg.getIndex());
	}
	
	@Override
	public String resolveNameXText(NameXPtg nameXPtg) {
		return index2name.get(nameXPtg.getNameIndex());
	}
	
	/**
	 * @return internal or external sheet.
	 */
	public ExternalSheet getAnyExternalSheet(int externSheetIndex) {
		return index2sheet.get(externSheetIndex);
	}
	
	@Override
	public ExternalSheet getExternalSheet(int externSheetIndex) {
		// return external sheet object if only if the sheet is exact external
		ExternalSheet externalSheet = getAnyExternalSheet(externSheetIndex);
		return externalSheet.getWorkbookName() != null ? externalSheet : null;
	}
	
	@Override
	public String getSheetNameByExternSheet(int externSheetIndex) {
		// get sheet no matter external or internal, and covert to 3D ref. if any
		ExternalSheet sheet = getAnyExternalSheet(externSheetIndex);
		String name = sheet.getSheetName();
		String lastName = sheet.getLastSheetName();
		return name.equals(lastName) ? name : sheet + ":" + lastName;
	}

	@Override
	public String getExternalLinkIndexFromBookName(String bookname) {
		return bookname;
	}

	/**
	 * name to represent named range
	 * @author Pao
	 */
	private class SimpleName implements EvaluationName {

		private final String name;
		private final int nameIndex;
		private int sheetIndex;

		/**
		 * @param sheetIndex sheet index; if -1, indicates whole book.
		 */
		public SimpleName(String name, int nameIndex, int sheetIndex) {
			this.name = name;
			this.nameIndex = nameIndex;
			this.sheetIndex = sheetIndex;
		}

		public Ptg[] getNameDefinition() {
			return FormulaParser.parse(name, ParsingBook.this, FormulaType.NAMEDRANGE, sheetIndex);
		}

		public String getNameText() {
			return name;
		}

		public boolean hasFormula() {
			return false;
		}

		public boolean isFunctionName() {
			return false;
		}

		public boolean isRange() {
			return false;
		}

		public NamePtg createPtg() {
			return new NamePtg(nameIndex);
		}
	}

}
