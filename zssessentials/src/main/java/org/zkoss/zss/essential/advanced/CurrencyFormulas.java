package org.zkoss.zss.essential.advanced;

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
