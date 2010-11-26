/* Formulas.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 25, 2010 7:02:49 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.formula;

import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

/**
 * @author Sam
 *
 */
public class Formulas {
	
	/* Category */
	private final static String FINANCIAL = "Financial";
	private final static String DATE_TIME = "Date & Time";
	private final static String Math_TRIG = "Math & Trig";
	private final static String STATISTICAL = "Statistical";
	private final static String LOOKUP_REF = "Lookup & Reference";
	private final static String DATABASE = "Database";
	private final static String TEXT = "Text";
	private final static String LOGICAL = "Logical";
	private final static String INFOMATION = "Information";
	private final static String ENGINEERING = "Engineering";
	
	//public final static int PARMETER_TYPE_DATE = 0;
	
	static LinkedHashMap<String, List<FormulaMetaInfo>> formulaInfos = 
		new LinkedHashMap<String, List<FormulaMetaInfo>>();
	
	//TODO use i18n
	static {
		/*Create Category List*/
		
		/* FINANCIAL */
		List<FormulaMetaInfo> financeAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(FINANCIAL, financeAry);
		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"ACCRINT", 
					"ACCRINT(issue, first_interest, settlement, rate, par, frequency, basis, calc_method)",
					"Returns the accrued interest for a security that pays periodic interest.",
					8, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"ACCRINTM",
					"ACCRINTM(issue, settlement, rate, par, basis)",
					"Returns the accrued interest for a security that pays interest at maturity,",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"AMORDEGRC",
					"AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, basis)",
					"Returns the prorated linear depreciation of an asset for each accounting period.",
					7, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"AMORLINC",
					"AMORLINC(cost, date_purchased, first_period, salvage, period, rate, basis)",
					"Returns the prorated linear depreciation of an asset for each accounting period.",
					7, false));

		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"COUPDAYBS",
					"COUPDAYBS(settlement, maturity, frequency, basis)",
					"Returns the number of days from the beginning of the coupon eriod to the settlement date.",
					3, false));

		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"COUPDAYS",
					"COUPDAYS(settlement, maturity, frequency, basis)",
					"Returns the number of days in the coupon period that contains the settlement date.",
					4, false));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPDAYSNC",
					"COUPDAYSNC(settlement, maturity, frequency, basis)",
					"Returns the number of days from the settlement date to the next coupon date.",
					4, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPNCD",
					"COUPNCD(settlement, maturity, frequency, basis)",
					"Returns the next coupon date after the settlement date.",
					4, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPNUM",
					"COUPNUM(settlement, maturity, frequency, basis)",
					"Returns the number of coupons payable between the settlement date and maturity date.",
					4, false));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPPCD",
					"COUPPCD(settlement, maturity, frequency, basis)",
					"Returns the previous coupon date before the settlement date.",
					4, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"CUMIPMT",
					"CUMIPMT(rate, nper, pv, start_period, end_period, type)",
					"Returns the cumulative interest paid between two periods.",
					6, false));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"CUMPRINC",
					"CUMPRINC(rate, nper, pv, start_period, end_period, type)",
					"Returns the cumulative principal paid on a load between two periods.",
					6, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DB",
					"DB(cost, salvage, life, period, month)",
					"Returns the depreciation of an asset for a specified period using the fixed-declining balance method",
					5, false));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DDB",
					"DDB(cost salvage, life, period, factor)",
					"Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DISC",
					"DISC(settlement, maturity, pr, redemption, basis)",
					"Returns the discount rate for a security",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DOLLARDE",
					"DOLLARDE(fractional_dollar, fraction)",
					"Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number",
					2, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DOLLARFR",
					"DOLLARFR(decimal_dollar, fraction)",
					"Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction",
					2, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DURATION",
					"DURATION(settlement, maturity, coupon, yld, frequency, basis)",
					"Returnsthe annual duration of a security with periodic interest payments.",
					6, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"EFFECT",
					"EFFECT(nominal_rate, npery)",
					"Returns the effective annual interest rate.",
					2, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"FV",
					"FV(rate, nper, pmt, pv, type)",
					"Returns the future value of an investment based on periodic, constant payments and a constant interset rate.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"FVSCHEDULE",
					"FVSCHEDULE(principal, sehedule)",
					"Returns the future value of an initial principal after applying a series of compound interest rates.",
					2, false));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"INTRATE",
					"INTRATE(settlement, maturity, investment, redemption, basis)",
					"Returns the interest rate for a fully invested security.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"IPMT",
					"IPMT(rate, per, nper, pv, fv, type)",
					"Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.",
					6, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"IRR",
					"IRR(values, guess)",
					"Returns the internal rate of return for a series of cash flows.",
					2, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ISPMT",
					"ISPMT(rate, per, nper, pv)",
					"Returns the interest paid during a specific period of an investment.",
					4, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"MIRR",
					"MIRR(values, finance_rate, reinvest_rate)",
					"Returns the modified internal rate of return for a series of periodic cash flows. MIRR considers both the cost of the investment and the interest received on reinvestment of cash.",
					3, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"MDURATION",
					"MDURATION(settlement, maturity, coupon, yld, frequency, basis)",
					"Returns the modified Macauley duration for a security with an assumed par value of $100.",
					6, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"NOMINAL",
					"NOMINAL(effect_rate,npery)",
					"Returns the nominal annual interest rate, given the effective rate and the number of compounding periods per year.",
					2, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"NPER",
					"NPER(rate, pmt, pv, fv, type)",
					"Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"NPV",
					"NPV(rate, value1, value2,...)",
					"Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).",
					3, true));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDFPRICE",
					"ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption,...)",
					"Returns the price per $100 face value of a security with an odd first period.",
					7, true));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDFYIELD",
					"ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption,...)",
					"Returns the yield of a security with an odd first period.",
					7, true));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDLPRICE",
					"ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency,...)",
					"Returns the price per $100 face value of a security with an odd last period.",
					8, true));
		
		/**
		 * TODO: check this formula, see how many parameter needs
		 */
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDLYIELD",
					"ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency,...)",
					"Returns the yield of a security with an odd last period.",
					7, true));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PMT",
					"PMT(rate, nper, pv, fv, type)",
					"Calculates the payment for a loan based on constant payments and a constant interest rate.",
					5, false));
					
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PPMT",
					"PPMT(rate, per, nper, pv, fv, type)",
					"Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.",
					6, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PRICE",
					"PRICE(settlement, maturity, rate, yld, redemption, frequency, basis)",
					"Returns the price per $100 face value of a security that pays periodic interest.",
					7, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PRICEDISC",
					"PRICEDISC(settlement, maturity, discount, redemption, basis)",
					"Returns the price per $100 face value of a discounted security.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PRICEMAT",
					"PRICEMAT(settlement, maturity, issue, rate, yld, basis)",
					"Returns the price per $100 face value of a security that pays interest at maturity.",
					6, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PV",
					"PV(rate, nper, pmt, fv, type)",
					"Returns the present value of an investment: the total amount that a series of future payments is worth now.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"RATE",
					"RATE(nper, pmt, pv, fv, type, guess)",
					"Returns the interest rate per period of a loan or an investment.For example, use 6%/4 for quartely payments at 6% APR.",
					6, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"RECEIVED",
					"RECEIVED(settlement, maturity, investment, discount, basis)",
					"Returns the amount received at maturity for a fully invested security.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"SLN",
					"SLN(cost, salvage, life)",
					"Returns the straight-line depreciation of an asset for one period.",
					3, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"SYD",
					"SYD(cost, salvage, life, per)",
					"Returns the sum-of years' digits depreciation of an asset for a specfied period.",
					4, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"TBILLEQ",
					"TBILLEQ(settlement, maturity, discount)",
					"Returns the bond-equivalent yield for a treasury bill.",
					3, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"TBILLPRICE",
					"TBILLPRICE(settlement, maturity, discount)",
					"Returns the price per $100 face value for a treasury bill.",
					3, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"TBILLYIELD",
					"TBILLYIELD(settlement, maturity ,pr)",
					"Returns the yield for a treasury bill",
					3, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"VDB",
					"VDB(cost, salvage, life, start_period, end_period, factor, no_switch)",
					"Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify",
					7, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"XIRR",
					"XIRR(values, dates, guess)",
					"Returns the internal rate of return for a schedule of cash flows.",
					3, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"XNPV",
					"XNPV(rate, values, dates)",
					"Returns the net present value for a schedule of case flows.",
					3, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"YIELD",
					"YIELD(settlement, maturity, rate, pr, redemption, frequency, basus)",
					"Returns the yield on a security that pays periodic interest.",
					7, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"YIELDDISC",
					"YIELDDISC(settlement, maturity, pr, redemption, basis)",
					"Returns the annual yield for a discounted security. For example, a treasury bill.",
					5, false));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"YIELDMAT",
					"YIELDMAT(settlement, maturity, issue, rate, pr, basis)",
					"Returns the annual yield of a security that pays interest at maturity",
					6, false));
		
		/* Statistical */
		List<FormulaMetaInfo> statAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(STATISTICAL, statAry);
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"AVEDEV",
					"AVEDEV(number1,number2,...)",
					"Returns the average of the absolute deviations of data points from their mean. AVEDEV is a measure of the variability in a data set.",
					2, true));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"AVERAGE",
					"AVERAGE(number1, [number2],...)",
					"Returns the average of its arguments",
					1, true));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"AVERAGEA",
					"AVERAGEA(value1,value2,...)",
					"Returns the average of its arguments, including numbers, text, and logical values",
					2, true));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"BINOMDIST",
					"BINOMDIST(number_s, trials, probability_s, cumulative)",
					"Returns the individual term binomial distribution probability",
					4, false));

		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"CHIDIST",
					"CHIDIST(x, degrees_freedom)",
					"Returns the one-tailed probability of the chi-squared distribution",
					2, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"CHIINV",
					"CHIINV(probability, degrees_freedom)",
					"Returns the inverse of the one-tailed probability of the chi-squared distribution",
					2, false));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"COUNT",
					"COUNT(value1, [value2],...)",
					"Counts how many numbers are in the list of arguments",
					1, true));

		statAry.add(new FormulaMetaInfo(STATISTICAL,	
					"COUNTA",
					"COUNTA(value1, [value2], ...)",
					"Counts how many values are in the list of arguments",
					1, true));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,	
					"EXPONDIST",
					"EXPONDIST(x, lambda, cumulative)",
					"Returns the exponential distribution",
					3, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,	
					"FDIST",
					"FDIST(x, degrees_freedom1, degrees_freedom2)",
					"Returns the F probability distribution",
					3, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"FINV",
					"FINV(probability, degrees_freedom1, degrees_freedom2)",
					"Returns the inverse of the F probability distribution",
					3, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GAMMADIST",
					"GAMMADIST(x, alpha, beta, cumulative)",
					"Returns the gamma distribution",
					4, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GAMMAINV",
					"GAMMAINV(probability, alpha, beta)",
					"Returns the inverse of the gamma cumulative distribution",
					3, false));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GAMMALN",
					"GAMMALN(x)",
					"Returns the natural logarithm of the gamma function",
					1, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GEOMEAN",
					"GEOMEAN(number1,number2,...)",
					"Returns the geometric mean",
					2, true));
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"HYPGEOMDIST",
					"HYPGEOMDIST(sample_s, number_sample, population_s, number_population)",
					"Returns the hypergeometric distribution",
					4, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"INTERCEPT",
					"INTERCEPT(known_y's, known_x's)",
					"Returns the intercept of the linear regression line",
					2, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"LOGEST",
					"LOGEST(known_y's, known_x's, const,stats)",
					"Returns statistics that describe an exponential curve matching known data points.",
					3, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"KURT",
					"KURT(number1, number2,...)",
					"Returns the kurtosis of a data set",
					2, true));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"MAX",
					"MAX(number1, number2,...)",
					"Returns the maximum value in a list of arguments",
					2, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"MIN",
					"MIN(number1, number2,...)",
					"Returns the minimum value in a list of arguments",
					2, true));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"NORMDIST",
					"NORMDIST(x, mean, standard_dev, cumulative)",
					"Returns the normal cumulative distribution",
					4, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"POISSON",
					"POISSON(x, mean, cumulative)",
					"Returns the Poisson distribution",
					3, false));

		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"SKEW",
					"SKEW(number1, number2,...)",
					"Returns the skewness of a distribution",
					2, false));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"SLOPE",
					"SLOPE(known_y's, known_x's)",
					"Returns the slope of the linear regression line",
					2, false));

		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"STDEV",
					"STDEV(number1, number2,...)",
					"Estimates standard deviation based on a sample",
					2, true));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"STDEVP",
					"STDEVP(number1, number2,...)",
					"Calculates standard deviation based on the entire population",
					2, true));
					
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"TDIST",
					"TDIST(x, degrees_freedom, tails)",
					"Returns the Student's t-distribution", 
					3, false));
					
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"TINV",
					"TINV(probability, degrees_freedom)",
					"Returns the inverse of the Student's t-distribution",
					2, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"VAR",
					"VAR(number1, number2,...)",
					"Estimates variance based on a sample",
					2, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"VARP",
					"VARP(number1, number2,...)",
					"Calculates variance based on the entire population",
					2, false));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"WEIBULL",
					"WEIBULL(x, alpha, beta, cumulative)",
					"Returns the Weibull distribution",
					4, false));
		
	}
	
	public static LinkedHashMap<String, List<FormulaMetaInfo>> getFormulaInfos() {
		return formulaInfos;
	}
}
