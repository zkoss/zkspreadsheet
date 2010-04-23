/* FinanceFunctionImpl.java

	Purpose:
		
	Description:
		
	History:
		Apr 19, 2010 5:52:10 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.formula.fn;

import java.math.BigDecimal;

import org.apache.poi.hssf.record.formula.eval.ErrorEval;
import org.apache.poi.hssf.record.formula.eval.EvaluationException;
import org.apache.poi.hssf.record.formula.eval.ValueEval;
import org.apache.poi.hssf.record.formula.functions.FinanceFunction;
import org.apache.poi.hssf.record.formula.functions.Function;
import org.apache.poi.hssf.record.formula.functions.NumericFunction;

/**
 * Implementation of Spreadsheet financial functions.
 * @author henrichen
 */
public class FinanceFunctionImpl {
	public static final Function DB = new NumericFunction() {
		@Override
		public double eval(ValueEval[] args, int srcRowIndex, int srcColumnIndex) throws EvaluationException {
			double cost = args[0] == null ? 0.0 : NumericFunction.singleOperandEvaluate(args[0], srcRowIndex, srcColumnIndex);
			if(cost < 0) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			}
			double salvage = args[1] == null ? 0.0 : NumericFunction.singleOperandEvaluate(args[1], srcRowIndex, srcColumnIndex);
			if(salvage < 0) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			}
			if(args[2] == null || args[3] == null) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			} else {
				double life = NumericFunction.singleOperandEvaluate(args[2], srcRowIndex, srcColumnIndex);
				double period = NumericFunction.singleOperandEvaluate(args[3], srcRowIndex, srcColumnIndex);
				if(life < 0 || period < 0) {
					throw new EvaluationException(ErrorEval.NUM_ERROR);
				}
				if(period-1 > life) {
					throw new EvaluationException(ErrorEval.NUM_ERROR);
				} else {
					double month = 12.0;
					if(args.length > 4) {
						month = NumericFunction.singleOperandEvaluate(args[4], srcRowIndex, srcColumnIndex);
					}
					double rate = new BigDecimal(1-Math.pow((salvage/cost), 1/life))
						.setScale(3, BigDecimal.ROUND_HALF_UP).doubleValue();
					return db(cost, salvage, life, period, month, rate);
				}
			}
			
		}
	};
	private static double db(double cost, double salvage, double life, double period, double month, double rate) {
		double db = cost*rate*month/12;
		double d = cost*rate*month/12;
		if(period>1) {
			for(int i = 1; i <= period-1; i++) {
				if(i<life) {
					db = (cost - d)*rate;
					d = d + db;
				} else {
					db = (cost - d)*rate*(12-month)/12;
				}
			}
		}
		return db;
	}
	
	public static final Function DDB = new NumericFunction() {
		@Override
		public double eval(ValueEval[] args, int srcRowIndex, int srcColumnIndex) throws EvaluationException {
			double cost = args[0] == null ? 0.0 : NumericFunction.singleOperandEvaluate(args[0], srcRowIndex, srcColumnIndex);
			if(cost < 0) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			}
			double salvage = args[1] == null ? 0.0 : NumericFunction.singleOperandEvaluate(args[1], srcRowIndex, srcColumnIndex);
			if(salvage < 0) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			}
			if(args[2] == null || args[3] == null) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			} else {
				double life = NumericFunction.singleOperandEvaluate(args[2], srcRowIndex, srcColumnIndex);
				double period = NumericFunction.singleOperandEvaluate(args[3], srcRowIndex, srcColumnIndex);
				if(life < 0 || period < 0) {
					throw new EvaluationException(ErrorEval.NUM_ERROR);
				}
				if(period > life) {
					throw new EvaluationException(ErrorEval.NUM_ERROR);
				} else {
					double factor = 2;
					if(args.length > 4) {
						factor = NumericFunction.singleOperandEvaluate(args[4], srcRowIndex, srcColumnIndex);
					}
					return ddb(cost, salvage, life, period, factor);
				}
			}
		}
	};
	private static double ddb(double cost, double salvage, double life, double period, double factor) {
		double d1 = factor*cost*Math.pow((life-factor)/life, period-1)/life;
		double d2 = (factor*cost*Math.pow((life-factor)/life, period-1)/life) - ((1-Math.pow((life-factor)/life, period))*cost - (cost-salvage));
		return Math.min(d1, d2);
	}

	public static final Function IPMT = new NumericFunction() {
		@Override
		public double eval(ValueEval[] args, int srcRowIndex, int srcColumnIndex) throws EvaluationException {
			double rate= NumericFunction.singleOperandEvaluate(args[0], srcRowIndex, srcColumnIndex);
			double per = NumericFunction.singleOperandEvaluate(args[1], srcRowIndex, srcColumnIndex);
			double nper = NumericFunction.singleOperandEvaluate(args[2], srcRowIndex, srcColumnIndex);
			if (per > nper) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			} else {
				double pv = NumericFunction.singleOperandEvaluate(args[3], srcRowIndex, srcColumnIndex);
				double fv = (args.length > 4 && args[4] != null) ?
						NumericFunction.singleOperandEvaluate(args[4], srcRowIndex, srcColumnIndex) : 0.0;
				int type = args.length > 5 ? 
						(int) NumericFunction.singleOperandEvaluate(args[5], srcRowIndex, srcColumnIndex) : 0;
				return ipmt(rate, per, nper, pv, fv, type);
			}
		}
	};
	
	public static final Function PPMT = new NumericFunction() {
		@Override
		public double eval(ValueEval[] args, int srcRowIndex, int srcColumnIndex) throws EvaluationException {
			double rate= NumericFunction.singleOperandEvaluate(args[0], srcRowIndex, srcColumnIndex);
			double per = NumericFunction.singleOperandEvaluate(args[1], srcRowIndex, srcColumnIndex);
			double nper = NumericFunction.singleOperandEvaluate(args[2], srcRowIndex, srcColumnIndex);
			if (per > nper) {
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			} else {
				double pv = NumericFunction.singleOperandEvaluate(args[3], srcRowIndex, srcColumnIndex);
				double fv = (args.length > 4 && args[4] != null) ?
						NumericFunction.singleOperandEvaluate(args[4], srcRowIndex, srcColumnIndex) : 0.0;
				int type = args.length > 5 ? 
						(int) NumericFunction.singleOperandEvaluate(args[5], srcRowIndex, srcColumnIndex) : 0;
				return ppmt(rate, per, nper, pv, fv, type);
			}
		}
	};
	private static double ppmt(double rate, double per, double nper, double pv, double fv, int type) {
		return pmt(rate, nper,pv, fv, type) - ipmt(rate, per, nper, pv, fv, type);
	}
	private static double pmt(double rate, double nper, double pv, double fv, int type) {
		double pmt = 0.0;
		if(rate == 0) {
			pmt = (fv + pv)/nper;
		} else {
			double term = Math.pow(1.0 + rate, nper);
			if (type == 1)
				pmt = (fv * rate / (term - 1.0) + pv * rate / (1.0 - 1.0 / term)) / (1.0+rate);
			else
			    pmt = fv * rate / (term - 1.0)  + pv * rate / (1.0 - 1.0 / term);
		}
		return -pmt;
	}
	private static double ipmt(double rate, double per, double nper, double pv, double fv, int type) {
		double ipmt = 0;
		
		// first period, different PayType has different value.
		if(per == 1) {
			ipmt =  type == 0 ?	-pv : 0.0;
		} else{
			double pmt = pmt(rate, nper, pv, fv, type);
			if(type == 0)
			{
				ipmt = GetZw(rate, per-1.0, pmt, pv, 0);
			}
			else if(type == 1)
			{
				ipmt = (GetZw(rate, per-2, pmt, pv, 1) - pmt);
			}
		}
		
		return ipmt * rate;

	}
	private static double GetZw(double rate, double nper, double pmt, double pv, int type) {
		double Zw = 0;
		if (rate == 0.0)
			Zw = pv + pmt * nper;
		else {
			double term = Math.pow(1.0 + rate, nper);
			if (type > 0.0)
				Zw = pv * term + pmt*(1.0 + rate)*(term - 1.0)/rate;
			else
				Zw = pv * term + pmt*(term - 1.0)/rate;
		}
		return -Zw;
	}
	
	public static final Function SLN = new FinanceFunction() {
		@Override
		protected double evaluate(double cost, double salvage, double life, double arg3, boolean type) throws EvaluationException {
			if (life == 0)
				throw new EvaluationException(ErrorEval.DIV_ZERO);
			return (cost-salvage)/life;
		}
	};
	
	public static final Function SYD = new FinanceFunction() {
		@Override
		protected double evaluate(double cost, double salvage, double life, double per, boolean type) throws EvaluationException {
			if(life <= 0 || per <= 0)
				throw new EvaluationException(ErrorEval.NUM_ERROR);
			return ((cost - salvage)*(life - per + 1)*2)/(life*(life + 1));
		}
	};
}
