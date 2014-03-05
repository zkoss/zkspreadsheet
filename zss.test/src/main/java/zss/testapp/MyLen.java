package zss.testapp;

import org.zkoss.poi.ss.formula.eval.ErrorEval;
import org.zkoss.poi.ss.formula.eval.EvaluationException;
import org.zkoss.poi.ss.formula.eval.StringEval;
import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.functions.TextFunction;

public class MyLen extends TextFunction {

	@Override
	protected ValueEval evaluateFunc(ValueEval[] args, int srcCellRow, int srcCellCol)
			throws EvaluationException {
		if (args[0] instanceof StringEval){
			String text = ((StringEval)args[0]).getStringValue();
			if (text.length() == 1){
				return new StringEval(text.length()+" character");
			}else{
				return new StringEval(text.length()+" characters");
			}
		}
		return ErrorEval.VALUE_INVALID;
	}

}
