package org.zkoss.zss.ngmodel.impl.sys;

import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Locale;

import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;
import org.zkoss.zss.ngmodel.util.Validations;

public class InputEngineImpl implements InputEngine{

	public InputResult parseInput(String editText,InputParseContext context){
		Validations.argNotNull(editText,context);
		InputResultImpl result = new InputResultImpl(editText);
		if(editText!=null && !"".equals(editText)){
			Object value;
			if(editText.startsWith("=")){
				value = editText.substring(1);
				result.setType(CellType.FORMULA);
				result.setValue(value);
				return result;
			}
			
			
			SimpleDateFormat sdf[] = new SimpleDateFormat[]{
				new SimpleDateFormat("yyyy/MM/dd")
			};
			//guess date, before number
			for (SimpleDateFormat f : sdf) {
				try {
					value = f.parse(editText);
					result.setType(CellType.DATE);
					result.setValue(value);
					return result;
				} catch (ParseException e) {
				}
			}
			
			//guess number
			DecimalFormat df = new DecimalFormat("#####.####");
			try {
				value = df.parse(editText);
				result.setType(CellType.NUMBER);
				result.setValue(value);
				return result;
			} catch (ParseException e) {}



			//guess string
			result.setType(CellType.STRING);
			result.setValue(editText);
		}
		return result;
	}
}
