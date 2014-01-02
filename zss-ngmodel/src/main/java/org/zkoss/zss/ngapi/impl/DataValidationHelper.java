package org.zkoss.zss.ngapi.impl;

import java.util.Date;

import org.zkoss.util.Locales;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NDataValidation.ValidationType;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.FormulaResultWrap;
import org.zkoss.zss.ngmodel.sys.CalendarUtil;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;

public class DataValidationHelper {

	private final NDataValidation validation;
	private final NSheet sheet;
	
	public DataValidationHelper(NDataValidation validation) {
		this.validation = validation;
		this.sheet = validation.getSheet();
	}

	public boolean validate(String editText, String dataformat) {
		final InputEngine ie = EngineFactory.getInstance().createInputEngine();
		final InputResult result = ie.parseInput(editText == null ? ""
				: editText, dataformat, new InputParseContext(Locales.getCurrent()));
		return validate(result.getType(),result.getValue());
	}
	
	public boolean validate(CellType cellType, Object value) {
		//allow any value => no need to do validation
		ValidationType vtype = validation.getValidationType();
		if (vtype == ValidationType.ANY) { //can be any value, meaning no validation
			return true;
		}
		//ignore empty and value is empty
		if (vtype!=ValidationType.TEXT_LENGTH 
				&& (value == null || (value instanceof String && ((String)value).length() == 0))) {
			if (validation.isEmptyCellAllowed()) {
				return true;
			}
		}
		//get new evaluated formula value 
		if (cellType == CellType.FORMULA) {
			
			FormulaEngine engine = EngineFactory.getInstance().createFormulaEngine();
			
			FormulaExpression expr = engine.parse((String)value, new FormulaParseContext(sheet, null));
			if(expr.hasError()){
				return false;
			}
			FormulaResultWrap result = new FormulaResultWrap(engine.evaluate(expr, new FormulaEvaluationContext(sheet)));
			
			value = result.getValue();
			cellType = result.getCellType();
		}
		CalendarUtil cal = EngineFactory.getInstance().getCalendarUtil();
		//start validation
		boolean success = true;
		switch(vtype) {
			// Integer ('Whole number') type
			case INTEGER:
				if (!isInteger(value) || !validateOperation((Number)value)) {
					success = false;
				}
				break;
			// Decimal type
			case DECIMAL:
				if (!isDecimal(value) || !validateOperation((Number)value)) {
					success = false;
				}
				break;
			// ZSS-260: the input value is a Date object, must convert it to Excel date type (a double number) before validating
			// Date type
			case DATE:
			// Time type
			case TIME:
				success = (value instanceof Date) && validateOperation(cal.dateToDoubleValue((Date)value));
				break;
			// List type ( combo box type )
			case LIST:
				if (!validateListOperation((value instanceof Date)?cal.dateToDoubleValue((Date)value):value)) {;
					success = false;
				}
				break;
			// String length type
			case TEXT_LENGTH:
				if (!isString(value) || !validateOperation(Integer.valueOf(value == null ? 0 : ((String)value).length()))) {
					success = false;
				}
				break;
			// Formula ( 'Custom' ) type
			case FORMULA:
				//zss 3.5 log it
				success = false;
//				throw new UnsupportedOperationException("Custom Validation is not supported yet!");
		}
		return success;
	}
	
	private static boolean isInteger(Object value) {
		if (value instanceof Number) {
			return ((Number)value).intValue() ==  ((Number)value).doubleValue();
		}
		return false;
	}
	private static boolean isDecimal(Object value) {
		return value instanceof Number;
	}
	
	private static boolean isString(Object value) {
		return value instanceof String; 
	}
	
	private boolean validateOperation(Number value) {
		if (value == null) {
			return false;
		}
		
		Object value1 = validation.getValue1(0);
		if (!(value1 instanceof Number)) { //type does not match
			return false;
		}
		Object value2 = validation.getValue2(0);
		double v1 = ((Number)value1).doubleValue();
		double v = value.doubleValue();
		double v2;
		switch(validation.getOperatorType()) {
		case BETWEEN:
			if (!(value2 instanceof Number)) { //type does not match
				return false;
			}
			v2 = ((Number)value2).doubleValue();
			return v >= v1 && v <= v2;
		case NOT_BETWEEN:
			if (!(value2 instanceof Number)) { //type does not match
				return false;
			}
			v2 = ((Number)value2).doubleValue();
			return v < v1 || v > v2;
		case EQUAL:
			return v == v1;
		case NOT_EQUAL:
			return v != v1;
		case GREATER_THAN:
				return v > v1;
		case LESS_THAN:
				return v < v1;
		case GREATER_OR_EQUAL:
				return v >= v1;
		case LESS_OR_EQUAL:
				return v <= v1;
		}
		return true;
	}
	
	private boolean validateListOperation(Object value) {
		if (value == null) {
			return false;
		}
		int size = validation.getNumOfValue1();
		for(int i=0;i<size;i++){
			Object val = validation.getValue1(i);
			if(value.equals(val)){
				return true;
			}
		}
		return false;
	}

}
