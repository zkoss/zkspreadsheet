package org.zkoss.zss.ngmodel.sys.input;

import java.util.Locale;

import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;


public interface InputEngine {

	public InputResult parseInput(String editText,String format, InputParseContext context);
}
