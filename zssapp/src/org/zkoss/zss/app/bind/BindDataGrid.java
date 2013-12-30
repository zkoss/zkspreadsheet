package org.zkoss.zss.app.bind;

import org.zkoss.zss.ngmodel.DefaultDataGrid;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;

public class BindDataGrid extends DefaultDataGrid {

	
	String bean1Value = "ABCD";
	double bean2Value = 33;
	
	private void updateBean1Value(String val){
		bean1Value = val;
		System.out.println("update bean1 value to "+val);
		//TODO how to notify bean1 change ?
	}
	
	private void updateBean2Value(Number val){
		bean2Value = val == null?0D:val.doubleValue();
		System.out.println("update bean2 value to "+val);
		//TODO how to notify bean1 change ?
	}
	
	@Override
	public NCellValue getValue(int rowIdx, int columnIdx) {
		NCellValue value = super.getValue(rowIdx, columnIdx);
		if(value!=null && value.getType() == CellType.FORMULA){
			FormulaExpression exp = (FormulaExpression)value.getValue();
			
			String expstr = exp.getFormulaString();
			if(!exp.hasError() && expstr.equals("Bean1")){
				return new NCellValue(bean1Value);
			}else if(!exp.hasError() && expstr.equals("Bean2")){
				return new NCellValue(bean2Value);
			}
			
		}
		return value;
	}

	@Override
	public void setValue(int rowIdx, int columnIdx, NCellValue value) {
		NCellValue oldvalue = super.getValue(rowIdx, columnIdx);
		if(oldvalue!=null && oldvalue.getType() == CellType.FORMULA){
			FormulaExpression exp = (FormulaExpression)oldvalue.getValue();
			String expstr = exp.getFormulaString();
			if(!exp.hasError() && expstr.equals("Bean1")){
				updateBean1Value(value==null?null:value.getType()==CellType.STRING?(String)value.getValue():"");
				return;
			}else if(!exp.hasError() && expstr.equals("Bean2")){
				updateBean2Value(value==null?0:value.getType()==CellType.NUMBER?(Number)value.getValue():0);
				return;
			}
		}
		super.setValue(rowIdx, columnIdx, value);
	}

	@Override
	public boolean validateValue(int rowIdx, int columnIdx, NCellValue value) {
		NCellValue oldvalue = super.getValue(rowIdx, columnIdx);
		if(oldvalue!=null && oldvalue.getType() == CellType.FORMULA){
			FormulaExpression exp = (FormulaExpression)oldvalue.getValue();
			String expstr = exp.getFormulaString();
			if(!exp.hasError() && expstr.equals("Bean1")){
				return value==null || value.getType()==CellType.STRING;
			}else if(!exp.hasError() && expstr.equals("Bean2")){
				return value==null || value.getType()==CellType.NUMBER;
			}
		}
		return super.validateValue(rowIdx, columnIdx, value);
	}

}
