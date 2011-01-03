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
	private final static String MATH_TRIG = "Math & Trig";
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
					8, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"ACCRINTM",
					"ACCRINTM(issue, settlement, rate, par, basis)",
					"Returns the accrued interest for a security that pays interest at maturity,",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"AMORDEGRC",
					"AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, basis)",
					"Returns the prorated linear depreciation of an asset for each accounting period.",
					7, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"AMORLINC",
					"AMORLINC(cost, date_purchased, first_period, salvage, period, rate, basis)",
					"Returns the prorated linear depreciation of an asset for each accounting period.",
					7, null));

		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"COUPDAYBS",
					"COUPDAYBS(settlement, maturity, frequency, basis)",
					"Returns the number of days from the beginning of the coupon eriod to the settlement date.",
					3, null));

		financeAry.add(new FormulaMetaInfo(FINANCIAL, 
					"COUPDAYS",
					"COUPDAYS(settlement, maturity, frequency, basis)",
					"Returns the number of days in the coupon period that contains the settlement date.",
					4, null));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPDAYSNC",
					"COUPDAYSNC(settlement, maturity, frequency, basis)",
					"Returns the number of days from the settlement date to the next coupon date.",
					4, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPNCD",
					"COUPNCD(settlement, maturity, frequency, basis)",
					"Returns the next coupon date after the settlement date.",
					4, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPNUM",
					"COUPNUM(settlement, maturity, frequency, basis)",
					"Returns the number of coupons payable between the settlement date and maturity date.",
					4, null));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"COUPPCD",
					"COUPPCD(settlement, maturity, frequency, basis)",
					"Returns the previous coupon date before the settlement date.",
					4, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"CUMIPMT",
					"CUMIPMT(rate, nper, pv, start_period, end_period, type)",
					"Returns the cumulative interest paid between two periods.",
					6, null));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"CUMPRINC",
					"CUMPRINC(rate, nper, pv, start_period, end_period, type)",
					"Returns the cumulative principal paid on a load between two periods.",
					6, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DB",
					"DB(cost, salvage, life, period, month)",
					"Returns the depreciation of an asset for a specified period using the fixed-declining balance method",
					5, null));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DDB",
					"DDB(cost salvage, life, period, factor)",
					"Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DISC",
					"DISC(settlement, maturity, pr, redemption, basis)",
					"Returns the discount rate for a security",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DOLLARDE",
					"DOLLARDE(fractional_dollar, fraction)",
					"Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number",
					2, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DOLLARFR",
					"DOLLARFR(decimal_dollar, fraction)",
					"Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction",
					2, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"DURATION",
					"DURATION(settlement, maturity, coupon, yld, frequency, basis)",
					"Returnsthe annual duration of a security with periodic interest payments.",
					6, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"EFFECT",
					"EFFECT(nominal_rate, npery)",
					"Returns the effective annual interest rate.",
					2, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"FV",
					"FV(rate, nper, pmt, pv, type)",
					"Returns the future value of an investment based on periodic, constant payments and a constant interset rate.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"FVSCHEDULE",
					"FVSCHEDULE(principal, sehedule)",
					"Returns the future value of an initial principal after applying a series of compound interest rates.",
					2, null));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"INTRATE",
					"INTRATE(settlement, maturity, investment, redemption, basis)",
					"Returns the interest rate for a fully invested security.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"IPMT",
					"IPMT(rate, per, nper, pv, fv, type)",
					"Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.",
					6, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"IRR",
					"IRR(values, guess)",
					"Returns the internal rate of return for a series of cash flows.",
					2, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ISPMT",
					"ISPMT(rate, per, nper, pv)",
					"Returns the interest paid during a specific period of an investment.",
					4, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"MIRR",
					"MIRR(values, finance_rate, reinvest_rate)",
					"Returns the modified internal rate of return for a series of periodic cash flows. MIRR considers both the cost of the investment and the interest received on reinvestment of cash.",
					3, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"MDURATION",
					"MDURATION(settlement, maturity, coupon, yld, frequency, basis)",
					"Returns the modified Macauley duration for a security with an assumed par value of $100.",
					6, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"NOMINAL",
					"NOMINAL(effect_rate, npery)",
					"Returns the nominal annual interest rate, given the effective rate and the number of compounding periods per year.",
					2, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"NPER",
					"NPER(rate, pmt, pv, fv, type)",
					"Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"NPV",
					"NPV(rate, value1, value2,...)",
					"Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).",
					3, "value"));

		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDFPRICE",
					"ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption, basis)",
					"Returns the price per $100 face value of a security with an odd first period.",
					7, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDFYIELD",
					"ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption, basis)",
					"Returns the yield of a security with an odd first period.",
					7, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDLPRICE",
					"ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency, basis)",
					"Returns the price per $100 face value of a security with an odd last period.",
					8, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"ODDLYIELD",
					"ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency, basis)",
					"Returns the yield of a security with an odd last period.",
					7, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PMT",
					"PMT(rate, nper, pv, fv, type)",
					"Calculates the payment for a loan based on constant payments and a constant interest rate.",
					5, null));
					
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PPMT",
					"PPMT(rate, per, nper, pv, fv, type)",
					"Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.",
					6, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PRICE",
					"PRICE(settlement, maturity, rate, yld, redemption, frequency, basis)",
					"Returns the price per $100 face value of a security that pays periodic interest.",
					7, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PRICEDISC",
					"PRICEDISC(settlement, maturity, discount, redemption, basis)",
					"Returns the price per $100 face value of a discounted security.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PRICEMAT",
					"PRICEMAT(settlement, maturity, issue, rate, yld, basis)",
					"Returns the price per $100 face value of a security that pays interest at maturity.",
					6, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"PV",
					"PV(rate, nper, pmt, fv, type)",
					"Returns the present value of an investment: the total amount that a series of future payments is worth now.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"RATE",
					"RATE(nper, pmt, pv, fv, type, guess)",
					"Returns the interest rate per period of a loan or an investment.For example, use 6%/4 for quartely payments at 6% APR.",
					6, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"RECEIVED",
					"RECEIVED(settlement, maturity, investment, discount, basis)",
					"Returns the amount received at maturity for a fully invested security.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"SLN",
					"SLN(cost, salvage, life)",
					"Returns the straight-line depreciation of an asset for one period.",
					3, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"SYD",
					"SYD(cost, salvage, life, per)",
					"Returns the sum-of years' digits depreciation of an asset for a specfied period.",
					4, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"TBILLEQ",
					"TBILLEQ(settlement, maturity, discount)",
					"Returns the bond-equivalent yield for a treasury bill.",
					3, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"TBILLPRICE",
					"TBILLPRICE(settlement, maturity, discount)",
					"Returns the price per $100 face value for a treasury bill.",
					3, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"TBILLYIELD",
					"TBILLYIELD(settlement, maturity ,pr)",
					"Returns the yield for a treasury bill",
					3, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"VDB",
					"VDB(cost, salvage, life, start_period, end_period, factor, no_switch)",
					"Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify",
					7, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"XIRR",
					"XIRR(values, dates, guess)",
					"Returns the internal rate of return for a schedule of cash flows.",
					3, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"XNPV",
					"XNPV(rate, values, dates)",
					"Returns the net present value for a schedule of case flows.",
					3, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"YIELD",
					"YIELD(settlement, maturity, rate, pr, redemption, frequency, basus)",
					"Returns the yield on a security that pays periodic interest.",
					7, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"YIELDDISC",
					"YIELDDISC(settlement, maturity, pr, redemption, basis)",
					"Returns the annual yield for a discounted security. For example, a treasury bill.",
					5, null));
		
		financeAry.add(new FormulaMetaInfo(FINANCIAL,
					"YIELDMAT",
					"YIELDMAT(settlement, maturity, issue, rate, pr, basis)",
					"Returns the annual yield of a security that pays interest at maturity",
					6, null));
		
		/* Statistical */
		List<FormulaMetaInfo> statAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(STATISTICAL, statAry);
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"AVEDEV",
					"AVEDEV(number1, number2,...)",
					"Returns the average of the absolute deviations of data points from their mean. AVEDEV is a measure of the variability in a data set.",
					2, "number"));
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"AVERAGE",
					"AVERAGE(number1, number2,...)",
					"Returns the average of its arguments",
					2, "number"));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"AVERAGEA",
					"AVERAGEA(value1, value2,...)",
					"Returns the average of its arguments, including numbers, text, and logical values",
					2, "value"));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"BINOMDIST",
					"BINOMDIST(number_s, trials, probability_s, cumulative)",
					"Returns the individual term binomial distribution probability",
					4, null));

		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"CHIDIST",
					"CHIDIST(x, degrees_freedom)",
					"Returns the one-tailed probability of the chi-squared distribution",
					2, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"CHIINV",
					"CHIINV(probability, degrees_freedom)",
					"Returns the inverse of the one-tailed probability of the chi-squared distribution",
					2, null));
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"COUNT",
					"COUNT(value1, value2,...)",
					"Counts how many numbers are in the list of arguments",
					1, "value"));
		statAry.add(new FormulaMetaInfo(STATISTICAL,	
					"COUNTA",
					"COUNTA(value1, value2, ...)",
					"Counts how many values are in the list of arguments",
					1, "value"));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,	
					"EXPONDIST",
					"EXPONDIST(x, lambda, cumulative)",
					"Returns the exponential distribution",
					3, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,	
					"FDIST",
					"FDIST(x, degrees_freedom1, degrees_freedom2)",
					"Returns the F probability distribution",
					3, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"FINV",
					"FINV(probability, degrees_freedom1, degrees_freedom2)",
					"Returns the inverse of the F probability distribution",
					3, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GAMMADIST",
					"GAMMADIST(x, alpha, beta, cumulative)",
					"Returns the gamma distribution",
					4, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GAMMAINV",
					"GAMMAINV(probability, alpha, beta)",
					"Returns the inverse of the gamma cumulative distribution",
					3, null));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GAMMALN",
					"GAMMALN(x)",
					"Returns the natural logarithm of the gamma function",
					1, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"GEOMEAN",
					"GEOMEAN(number1, number2,...)",
					"Returns the geometric mean",
					2, "number"));
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"HYPGEOMDIST",
					"HYPGEOMDIST(sample_s, number_sample, population_s, number_population)",
					"Returns the hypergeometric distribution",
					4, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"INTERCEPT",
					"INTERCEPT(known_y's, known_x's)",
					"Returns the intercept of the linear regression line",
					2, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"LOGEST",
					"LOGEST(known_y's, known_x's, const,stats)",
					"Returns statistics that describe an exponential curve matching known data points.",
					3, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"KURT",
					"KURT(number1, number2,...)",
					"Returns the kurtosis of a data set",
					2, "number"));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"MAX",
					"MAX(number1, number2,...)",
					"Returns the maximum value in a list of arguments",
					2, "number"));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"MIN",
					"MIN(number1, number2,...)",
					"Returns the minimum value in a list of arguments",
					2, "number"));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"NORMDIST",
					"NORMDIST(x, mean, standard_dev, cumulative)",
					"Returns the normal cumulative distribution",
					4, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"POISSON",
					"POISSON(x, mean, cumulative)",
					"Returns the Poisson distribution",
					3, null));

		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"SKEW",
					"SKEW(number1, number2,...)",
					"Returns the skewness of a distribution",
					2, "number"));
	
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"SLOPE",
					"SLOPE(known_y's, known_x's)",
					"Returns the slope of the linear regression line",
					2, null));

		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"STDEV",
					"STDEV(number1, number2,...)",
					"Estimates standard deviation based on a sample",
					2, "number"));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"STDEVP",
					"STDEVP(number1, number2,...)",
					"Calculates standard deviation based on the entire population",
					2, "number"));
					
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"TDIST",
					"TDIST(x, degrees_freedom, tails)",
					"Returns the Student's t-distribution", 
					3, null));
					
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"TINV",
					"TINV(probability, degrees_freedom)",
					"Returns the inverse of the Student's t-distribution",
					2, null));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"VAR",
					"VAR(number1, number2,...)",
					"Estimates variance based on a sample",
					2, "number"));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"VARP",
					"VARP(number1, number2,...)",
					"Calculates variance based on the entire population",
					2, "number"));
		
		statAry.add(new FormulaMetaInfo(STATISTICAL,
					"WEIBULL",
					"WEIBULL(x, alpha, beta, cumulative)",
					"Returns the Weibull distribution",
					4, null));
		/* Math */
		List<FormulaMetaInfo> mathAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(MATH_TRIG, mathAry);
		
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ABS",
					"ABS(number)",
					"Returns the absolute value of a number",
					4, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ACOS",
					"ACOS(number)",
					"Returns the arccosine of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ACOSH",
					"ACOSH(number)",
					"Returns the inverse hyperbolic cosine of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ASIN",
					"ASIN(number)",
					"Returns the arcsine of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ASINH",
					"ASINH(number)",
					"Returns the inverse hyperbolic sine of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ATAN",
					"ATAN(number)",
					"Returns the arctangent of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ATAN2",
					"ATAN2(x_num, y_num)",
					"Returns the arctangent from x- and y-coordinates",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ATANH",
					"ATANH(number)",
					"Returns the inverse hyperbolic tangent of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"CEILING",
					"CEILING(number, significance)",
					"Rounds a number to the nearest integer or to the nearest multiple of significance",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"COMBIN",
					"COMBIN(number, number_chosen)",
					"Returns the number of combinations for a given number of objects",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"COS",
					"COS(number)",
					"Returns the cosine of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"COSH",
					"COSH(number)",
					"Returns the hyperbolic cosine of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"DEGREES",
					"DEGREES(angle)",
					"Converts radians to degrees",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"EVEN",
					"EVEN(number)",
					"Rounds a number up to the nearest even integer",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"EXP",
					"EXP(number)",
					"Returns e raised to the power of a given number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"FACT",
					"FACT(number)",
					"Returns the factorial of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"FACTDOUBLE",
					"FACTDOUBLE(number)",
					"Returns the double factorial of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"FLOOR",
					"FLOOR(number, significance)",
					"Rounds a number down, toward zero",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"GCD",
					"GCD(number1, number2, ...)",
					"Returns the greatest common divisor",
					2, "number"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"INT",
					"INT(number)",
					"Rounds a number down to the nearest integer",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"LCM",
					"LCM(number1, number2, ...)",
					"Returns the least common multiple",
					2, "number"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"LN",
					"LN(number)",
					"Returns the natural logarithm of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"LOG",
					"LOG(number, base)",
					"Returns the logarithm of a number to a specified base",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"LOG10",
					"LOG10(number)",
					"Returns the base-10 logarithm of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"MDETERM",
					"MDETERM(array)",
					"Returns the matrix determinant of an array",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"MINVERSE",
					"MINVERSE(array)",
					"Returns the matrix inverse of an array",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"MMULT",
					"MMULT(array1, array2)",
					"Returns the matrix product of two arrays",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"MOD",
					"MOD(number, divisor)",
					"Returns the remainder from division",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"MROUND",
					"MROUND(number, multiple)",
					"Returns a number rounded to the desired multiple",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"MULTINOMIAL",
					"MULTINOMIAL(number1, number2, ...)",
					"Returns the multinomial of a set of numbers",
					2, "number"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ODD",
					"ODD(number)",
					"Rounds a number up to the nearest odd integer",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"PI",
					"PI( )",
					"Returns the value of pi",
					0, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"POWER",
					"POWER(number, power)",
					"Returns the result of a number raised to a power",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"PRODUCT",
					"PRODUCT(number1, number2, ...)",
					"Multiplies its arguments",
					1, "number"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"QUOTIENT",
					"QUOTIENT(numerator, denominator)",
					"Returns the integer portion of a division",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"RADIANS",
					"RADIANS(angle)",
					"Converts degrees to radians",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"RAND",
					"RAND( )",
					"Returns a random number between 0 and 1",
					0, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"RANDBETWEEN",
					"RANDBETWEEN(bottom, top)",
					"Returns a random number between the numbers you specify",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ROMAN",
					"ROMAN(number, form)",
					"Converts an arabic numeral to roman, as text",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ROUND",
					"ROUND(number, num_digits)",
					"Rounds a number to a specified number of digits",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ROUNDDOWN",
					"ROUNDDOWN(number, num_digits)",
					"Rounds a number down, toward zero",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"ROUNDUP",
					"ROUNDUP(number,num_digits)",
					"Rounds a number up, away from zero",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,	
					"SERIESSUM",
					"SERIESSUM(x, n, m, coefficients)",
					"Returns the sum of a power series based on the formula",
					4, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,	
					"SIGN",
					"SIGN(number)",
					"Returns the sign of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,	
					"SIN",
					"SIN(number)",
					"Returns the sine of the given angle",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,	
					"SINH",
					"SINH(number)",
					"Returns the hyperbolic sine of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,	
					"SQRT",
					"SQRT(number)",
					"Returns a positive square root",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SQRTPI",
					"SQRTPI(number)",
					"Returns the square root of (number * pi)",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUBTOTAL",
					"SUBTOTAL(function_num, ref1, ref2, ...)",
					"Returns a subtotal in a list or database",
					3, "ref"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUM",
					"SUM(number1, number2, number3, ...)",
					"Adds its arguments",
					2, "number"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUMIF",
					"SUMIF(range, criteria, sum_range)",
					"Adds the cells specified by a given criteria",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUMPRODUCT",
					"SUMPRODUCT(array1, array2, array3, ...)",
					"Returns the sum of the products of corresponding array components",
					3, "array"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUMSQ",
					"SUMSQ(number1, number2, ...)",
					"Returns the sum of the squares of the arguments",
					2, "number"));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUMX2MY2",
					"SUMX2MY2(array_x, array_y)",
					"Returns the sum of the difference of squares of corresponding values in two arrays",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUMX2PY2",
					"SUMX2PY2(array_x, array_y)",
					"Returns the sum of the sum of squares of corresponding values in two arrays",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"SUMXMY2",
					"SUMXMY2(array_x, array_y)",
					"Returns the sum of squares of differences of corresponding values in two arrays",
					2, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"TAN",
					"TAN(number)",
					"Returns the tangent of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"TANH",
					"TANH(number)",
					"Returns the hyperbolic tangent of a number",
					1, null));
		mathAry.add(new FormulaMetaInfo(MATH_TRIG,
					"TRUNC",
					"TRUNC(number, num_digits)",
					"Truncates a number to an integer",
					2, null));

		/* Date and Time */
		List<FormulaMetaInfo> dateAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(DATE_TIME, dateAry);	
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"DATE",
					"DATE(year, month, day)",
					"Returns the serial number of a particular date",
					3, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"DATEVALUE",
					"DATEVALUE(date_text)",
					"Converts a date in the form of text to a serial number",
					1, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"DAY",
					"DAY(serial_number)",
					"Converts a serial number to a day of the month",
					1, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"DAYS360",
					"DAYS360(start_date, end_date, method)",
					"Calculates the number of days between two dates based on a 360-day year",
					2, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"HOUR",
					"HOUR(serial_number)",
					"Converts a serial number to an hour",
					1, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"MINUTE",
					"MINUTE(serial_number)",
					"Converts a serial number to a minute",
					1, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"MONTH",
					"MONTH(serial_number)",
					"Converts a serial number to a month",
					1, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"NOW",
					"NOW()",
					"Returns the serial number of the current date and time",
					0, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"SECOND",
					"SECOND(serial_number)",
					"Converts a serial number to a second",
					1, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"TIME",
					"TIME(hour, minute, second)",
					"Returns the serial number of a particular time",
					3, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"TODAY",
					"TODAY( )",
					"Returns the serial number of today's date",
					0, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"WEEKDAY",
					"WEEKDAY(serial_number, return_type)",
					"Converts a serial number to a day of the week",
					2, null));
		dateAry.add(new FormulaMetaInfo(DATE_TIME,
					"YEAR",
					"YEAR(serial_number)",
					"Converts a serial number to a year",
					1, null));
		
		/* Text */
		List<FormulaMetaInfo> textAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(TEXT, textAry);	
		textAry.add(new FormulaMetaInfo(TEXT,
					"ASC",
					"ASC(text)",
					"Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"BAHTTEXT",
					"BAHTTEXT(number)",
					"Converts a number to text, using the ß (baht) currency format",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"CHAR",
					"CHAR(number)",
					"Returns the character specified by the code number",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"CLEAN",
					"CLEAN(text)",
					"Removes all nonprintable characters from text",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"CODE",
					"CODE(text)",
					"Returns a numeric code for the first character in a text string",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"CONCATENATE",
					"CONCATENATE(text1, text2, ...)",
					"Joins several text items into one text item",
					1, "text"));
		textAry.add(new FormulaMetaInfo(TEXT,
					"DOLLAR",
					"DOLLAR(number, decimals)",
					"Converts a number to text, using the $ (dollar) currency format",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"EXACT",
					"EXACT(text1, text2)",
					"Checks to see if two text values are identical",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"FIND",
					"FIND(find_text, within_text, start_num)",
					"Finds one text value within another (case-sensitive)",
					3, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"FINDB",
					"FINDB(find_text, within_text, start_num)",
					"Finds one text value within another (case-sensitive)",
					3, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"FIXED",
					"FIXED(number, decimals, no_commas)",
					"Formats a number as text with a fixed number of decimals",
					3, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"JIS",
					"JIS(text)",
					"Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"LEFT",
					"LEFT(text, num_chars)",
					"Returns the leftmost characters from a text value",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"LEFTB",
					"LEFTB(text, num_bytes)",
					"Returns the leftmost characters from a text value",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"LEN",
					"LEN(text)",
					"Returns the number of characters in a text string",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"LENB",
					"LENB(text)",
					"Returns the number of bytes used to represent the characters in a text string.",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"LOWER",
					"LOWER(text)",
					"Converts text to lowercase",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"MID",
					"MID(text, start_num, num_chars)",
					"Returns a specific number of characters from a text string starting at the position you specify",
					3, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"MIDB",
					"MIDB(text,start_num,num_bytes)",
					"Returns a specific number of characters from a text string starting at the position you specify",
					3, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"PHONETIC",
					"PHONETIC(reference)",
					"Extracts the phonetic (furigana) characters from a text string",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"PROPER",
					"PROPER(text)",
					"Capitalizes the first letter in each word of a text value",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"REPLACE",
					"REPLACE(old_text, start_num, num_chars, new_text)",
					"Replaces characters within text",
					4, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"REPLACEB",
					"REPLACEB(old_text,start_num,num_bytes,new_text)",
					"Replaces characters within text",
					4, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"REPT",
					"REPT(text, number_times)",
					"Repeats text a given number of times",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"RIGHT",
					"RIGHT(text, num_chars)",
					"Returns the rightmost characters from a text value",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"RIGHTB",
					"RIGHTB(text,num_bytes)",
					"Returns the rightmost characters from a text value",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"SEARCH",
					"SEARCH(find_text, within_text, [start_num])",
					"Finds one text value within another (not case-sensitive)",
					3, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"SEARCHB",
					"SEARCHB(find_text, within_text, [start_num])",
					"Finds one text value within another (not case-sensitive)",
					3, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"SUBSTITUTE",
					"SUBSTITUTE(text, old_text, new_text, instance_num)",
					"Substitutes new text for old text in a text string",
					4, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"T",
					"T(value)",
					"Converts its arguments to text",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"TEXT",
					"TEXT(value, format_text)",
					"Formats a number and converts it to text",
					2, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"TRIM",
					"TRIM(text)",
					"Removes spaces from text",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"UPPER",
					"UPPER(text)",
					"Converts text to uppercase",
					1, null));
		textAry.add(new FormulaMetaInfo(TEXT,
					"VALUE",
					"VALUE(text)",
					"Converts a text argument to a number",
					1, null));
		
					
		/* Logical */
		List<FormulaMetaInfo> logAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(LOGICAL, logAry);
		logAry.add(new FormulaMetaInfo(LOGICAL,
				"AND",
				"AND(logical1, logical2, ...)", 
				"Returns TRUE if all of its arguments are TRUE", 
				2, "logical"));
		logAry.add(new FormulaMetaInfo(LOGICAL,
				"FALSE",
				"FALSE( )",
				"Returns the logical value FALSE",
				0, null));
		logAry.add(new FormulaMetaInfo(LOGICAL,
				"IF",
				"IF(logical_test, value_if_true, [value_if_false])",
				"Specifies a logical test to perform",
				3, null));
		logAry.add(new FormulaMetaInfo(LOGICAL,
				"NOT",
				"NOT(logical)",
				"Reverses the logic of its argument",
				1, null));
		logAry.add(new FormulaMetaInfo(LOGICAL,
				"OR",
				"OR(logical1, logical2,...)",
				"Returns TRUE if any argument is TRUE",
				2, "logical"));
		logAry.add(new FormulaMetaInfo(LOGICAL,
				"TRUE",
				"TRUE( )",
				"Returns the logical value TRUE",
				0, null));
		
		/* Infomation */
		List<FormulaMetaInfo> infoAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(INFOMATION, infoAry);
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"ISBLANK",
				"ISBLANK(value)",
				"Returns TRUE if the value is blank",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"ISLOGICAL",
				"ISLOGICAL(value)",
				"Returns TRUE if the value is a logical value",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"ISEVEN",
				"ISEVEN(number)",
				"Returns TRUE if the number is even",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"ISODD",
				"ISODD(number)",
				"Returns TRUE if the number is odd",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"N",
				"N(value)",
				"Returns a value converted to a number",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"NA",
				"NA( )",
				"Returns the error value #N/A",
				0, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"ISNUMBER",
				"ISNUMBER(value)",
				"Returns TRUE if the value is a number",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,
				"ISTEXT",
				"ISTEXT(value)",
				"Returns TRUE if the value is text",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,		
				"ISNONTEXT",
				"ISNONTEXT(value)",
				"Returns TRUE if the value is not text",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,	
				"ISERR",
				"ISERR(value)",
				"Returns TRUE if the value is any error value except #N/A",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,	
				"ISERROR",
				"ISERROR(value)",
				"Returns TRUE if the value is any error value",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,	
				"ISNA",
				"ISNA(value)",
				"Returns TRUE if the value is the #N/A error value",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,	
				"TYPE",
				"TYPE(value)",
				"Returns a number indicating the data type of a value",
				1, null));
		infoAry.add(new FormulaMetaInfo(INFOMATION,	
				"ERROR.TYPE",
				"ERROR.TYPE(error_val)",
				"Returns a number corresponding to an error type",
				1, null));
		
		/* Engineering */
		List<FormulaMetaInfo> engAry = new LinkedList<FormulaMetaInfo>();
		formulaInfos.put(ENGINEERING, engAry);
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"BESSELI",
				"BESSELI(x, n)",
				"Returns the modified Bessel function In(x)",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"BESSELJ",
				"BESSELJ(x, n)",
				"Returns the Bessel function Jn(x)",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"BESSELK",
				"BESSELK(x, n)",
				"Returns the modified Bessel function Kn(x)",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"BESSELY",
				"BESSELY(x, n)",
				"Returns the Bessel function Yn(x)",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"BIN2DEC",
				"BIN2DEC(number)",
				"Converts a binary number to decimal",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"BIN2HEX",
				"BIN2HEX(number, places)",
				"Converts a binary number to hexadecimal",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"BIN2OCT",
				"BIN2OCT(number, places)",
				"Converts a binary number to octal",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"COMPLEX",
				"COMPLEX(real_num, i_num, suffix)",
				"Converts real and imaginary coefficients into a complex number",
				3, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"CONVERT",
				"CONVERT(number, from_unit, to_unit)",
				"Converts a number from one measurement system to another",
				3, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"DEC2BIN",
				"DEC2BIN(number, places)",
				"Converts a decimal number to binary",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"DEC2HEX",
				"DEC2HEX(number, places)",
				"Converts a decimal number to hexadecimal",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"DEC2OCT",
				"DEC2OCT(number, places)",
				"Converts a decimal number to octal",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"DELTA",
				"DELTA(number1, number2)",
				"Tests whether two values are equal",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"ERF",
				"ERF(lower_limit, upper_limit)",
				"Returns the error function",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"ERFC",
				"ERFC(x)",
				"Returns the complementary error function",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"GESTEP",
				"GESTEP(number, step)",
				"Tests whether a number is greater than a threshold value",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"HEX2BIN",
				"HEX2BIN(number, places)",
				"Converts a hexadecimal number to binary",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"HEX2DEC",
				"HEX2DEC(number)",
				"Converts a hexadecimal number to decimal",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"HEX2OCT",
				"HEX2OCT(number, places)",
				"Converts a hexadecimal number to octal",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMABS",
				"IMABS(inumber)",
				"Returns the absolute value (modulus) of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMAGINARY",
				"IMAGINARY(inumber)",
				"Returns the imaginary coefficient of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMARGUMENT",
				"IMARGUMENT(inumber)",
				"Returns the argument theta, an angle expressed in radians",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMCONJUGATE",
				"IMCONJUGATE(inumber)",
				"Returns the complex conjugate of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMCOS",
				"IMCOS(inumber)",
				"Returns the cosine of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMDIV",
				"IMDIV(inumber1, inumber2)",
				"Returns the quotient of two complex numbers",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMEXP",
				"IMEXP(inumber)",
				"Returns the exponential of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMLN",
				"IMLN(inumber)",
				"Returns the natural logarithm of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMLOG10",
				"IMLOG10(inumber)",
				"Returns the base-10 logarithm of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMLOG2",
				"IMLOG2(inumber)",
				"Returns the base-2 logarithm of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMPOWER",
				"IMPOWER(inumber, number)",
				"Returns a complex number raised to an integer power",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMPRODUCT",
				"IMPRODUCT(inumber1, inumber2,...)",
				"Returns the product of complex numbers",
				2, "inumber"));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMREAL",
				"IMREAL(inumber)",
				"Returns the real coefficient of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMSIN",
				"IMSIN(inumber)",
				"Returns the sine of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMSQRT",
				"IMSQRT(inumber)",
				"Returns the square root of a complex number",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMSUB",
				"IMSUB(inumber1, inumber2)",
				"Returns the difference between two complex numbers",
				2, "inumber"));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"IMSUM",
				"IMSUM(inumber1,inumber2,...)",
				"Returns the sum of complex numbers",
				2, "inumber"));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"OCT2BIN",
				"OCT2BIN(number, places)",
				"Converts an octal number to binary",
				2, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"OCT2DEC",
				"OCT2DEC(number)",
				"Converts an octal number to decimal",
				1, null));
		engAry.add(new FormulaMetaInfo(ENGINEERING,
				"OCT2HEX",
				"OCT2HEX(number, places)",
				"Converts an octal number to hexadecimal",
				2, null));
	}
	
	public static LinkedHashMap<String, List<FormulaMetaInfo>> getFormulaInfos() {
		return formulaInfos;
	}
}
