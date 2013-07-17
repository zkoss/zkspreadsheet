package org.zkoss.zss.essential.advanced;

import org.zkoss.poi.ss.formula.eval.ValueEval;
import org.zkoss.poi.ss.formula.functions.AggregateFunction;

/**
 * Convert USD to TWD upon exchange rate.
 * @author Hawk
 *
 */
public class CurrencyFormulas {
	public static double usd2Twd(double usd, double exchangeRate) {
        return usd * exchangeRate;
    }
	
}
