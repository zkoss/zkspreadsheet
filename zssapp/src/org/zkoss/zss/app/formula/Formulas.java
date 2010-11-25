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
	
	static LinkedHashMap<String, FormulaMetaInfo> formulaInfos = new LinkedHashMap<String, FormulaMetaInfo>();
	
	//TODO use i18n
	static {
		/* FINANCIAL */
		formulaInfos.put(FINANCIAL, 
			new FormulaMetaInfo(FINANCIAL, 
					"ACCRINT", 
					"ACCRINT(issue, first_interest, settlement, rate, par, frequency, basis, calc_method)",
					"Returns the accrued interest for a security that pays periodic interest.",
					8));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL, 
					"ACCRINTM",
					"ACCRINTM(issue, settlement, rate, par, basis)",
					"Returns the accrued interest for a security that pays interest at maturity,",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL, 
					"AMORDEGRC",
					"AMORDEGRC(cost, date_purchased, first_period, salvage, period, rate, basis)",
					"Returns the prorated linear depreciation of an asset for each accounting period.",
					7));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL, 
					"AMORLINC",
					"AMORLINC(cost, date_purchased, first_period, salvage, period, rate, basis)",
					"Returns the prorated linear depreciation of an asset for each accounting period.",
					7));

		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL, 
					"COUPDAYBS",
					"COUPDAYBS(settlement, maturity, frequency, basis)",
					"Returns the number of days from the beginning of the coupon eriod to the settlement date.",
					3));

		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL, 
					"COUPDAYS",
					"COUPDAYS(settlement, maturity, frequency, basis)",
					"Returns the number of days in the coupon period that contains the settlement date.",
					4));

		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"COUPDAYSNC",
					"COUPDAYSNC(settlement, maturity, frequency, basis)",
					"Returns the number of days from the settlement date to the next coupon date.",
					4));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"COUPNCD",
					"COUPNCD(settlement, maturity, frequency, basis)",
					"Returns the next coupon date after the settlement date.",
					4));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"COUPNUM",
					"COUPNUM(settlement, maturity, frequency, basis)",
					"Returns the number of coupons payable between the settlement date and maturity date.",
					4));

		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"COUPPCD",
					"COUPPCD(settlement, maturity, frequency, basis)",
					"Returns the previous coupon date before the settlement date.",
					4));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"CUMIPMT",
					"CUMIPMT(rate, nper, pv, start_period, end_period, type)",
					"Returns the cumulative interest paid between two periods.",
					6));

		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"CUMPRINC",
					"CUMPRINC(rate, nper, pv, start_period, end_period, type)",
					"Returns the cumulative principal paid on a load between two periods.",
					6));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"DB",
					"DB(cost, salvage, life, period, month)",
					"Returns the depreciation of an asset for a specified period using the fixed-declining balance method",
					5));

		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"DDB",
					"DDB(cost salvage, life, period, factor)",
					"Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"DISC",
					"DISC(settlement, maturity, pr, redemption, basis)",
					"Returns the discount rate for a security",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"DOLLARDE",
					"DOLLARDE(fractional_dollar, fraction)",
					"Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number",
					2));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"DOLLARFR",
					"DOLLARFR(decimal_dollar, fraction)",
					"Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction",
					2));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"DURATION",
					"DURATION(settlement, maturity, coupon, yld, frequency, basis)",
					"Returnsthe annual duration of a security with periodic interest payments.",
					6));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"EFFECT",
					"EFFECT(nominal_rate, npery)",
					"Returns the effective annual interest rate.",
					2));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"FV",
					"FV(rate, nper, pmt, pv, type)",
					"Returns the future value of an investment based on periodic, constant payments and a constant interset rate.",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"FVSCHEDULE",
					"FVSCHEDULE(principal, sehedule)",
					"Returns the future value of an initial principal after applying a series of compound interest rates.",
					2));

		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"INTRATE",
					"INTRATE(settlement, maturity, investment, redemption, basis)",
					"Returns the interest rate for a fully invested security.",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"IPMT",
					"IPMT(rate, per, nper, pv, fv, type)",
					"Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate.",
					6));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"IRR",
					"IRR(values, guess)",
					"Returns the internal rate of return for a series of cash flows.",
					2));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"ISPMT",
					"ISPMT(rate, per, nper, pv)",
					"Returns the interest paid during a specific period of an investment.",
					4));
	
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"NPV",
					"NPV(rate, value1, value2,...)",
					"Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).",
					3));
		/**
		 * TODO: check this formula, see how many parameter needs
		 */
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"ODDFPRICE",
					"ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld, redemption,...)",
					"Returns the price per $100 face value of a security with an odd first period.",
					7));
		
		/**
		 * TODO: check this formula, see how many parameter needs
		 */
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"ODDFYIELD",
					"ODDFYIELD(settlement, maturity, issue, first_coupon, rate, pr, redemption,...)",
					"Returns the yield of a security with an odd first period.",
					7));
		
		/**
		 * TODO: check this formula, see how many parameter needs
		 */
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"ODDLPRICE",
					"ODDLPRICE(settlement, maturity, last_interest, rate, yld, redemption, frequency,...)",
					"Returns the price per $100 face value of a security with an odd last period.",
					8));
		
		/**
		 * TODO: check this formula, see how many parameter needs
		 */
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"ODDLYIELD",
					"ODDLYIELD(settlement, maturity, last_interest, rate, pr, redemption, frequency,...)",
					"Returns the yield of a security with an odd last period.",
					7));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"PMT",
					"PMT(rate, nper, pv, fv, type)",
					"Calculates the payment for a loan based on constant payments and a constant interest rate.",
					5));
					
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"PPMT",
					"PPMT(rate, per, nper, pv, fv, type)",
					"Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate.",
					6));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"PRICE",
					"PRICE(settlement, maturity, rate, yld, redemption, frequency, basis)",
					"Returns the price per $100 face value of a security that pays periodic interest.",
					7));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"PRICEDISC",
					"PRICEDISC(settlement, maturity, discount, redemption, basis)",
					"Returns the price per $100 face value of a discounted security.",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"PRICEMAT",
					"PRICEMAT(settlement, maturity, issue, rate, yld, basis)",
					"Returns the price per $100 face value of a security that pays interest at maturity.",
					6));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"PV",
					"PV(rate, nper, pmt, fv, type)",
					"Returns the present value of an investment: the total amount that a series of future payments is worth now.",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"RATE",
					"RATE(nper, pmt, pv, fv, type, guess)",
					"Returns the interest rate per period of a loan or an investment.For example, use 6%/4 for quartely payments at 6% APR.",
					6));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"RECEIVED",
					"RECEIVED(settlement, maturity, investment, discount, basis)",
					"Returns the amount received at maturity for a fully invested security.",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"SLN",
					"SLN(cost, salvage, life)",
					"Returns the straight-line depreciation of an asset for one period.",
					3));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"SYD",
					"SYD(cost, salvage, life, per)",
					"Returns the sum-of years' digits depreciation of an asset for a specfied period.",
					4));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"TBILLEQ",
					"TBILLEQ(settlement, maturity, discount)",
					"Returns the bond-equivalent yield for a treasury bill.",
					3));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"TBILLPRICE",
					"TBILLPRICE(settlement, maturity, discount)",
					"Returns the price per $100 face value for a treasury bill.",
					3));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"TBILLYIELD",
					"TBILLYIELD(settlement, maturity ,pr)",
					"Returns the yield for a treasury bill",
					3));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"VDB",
					"VDB(cost, salvage, life, start_period, end_period, factor, no_switch)",
					"Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify",
					7));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"XIRR",
					"XIRR(values, dates, guess)",
					"Returns the internal rate of return for a schedule of cash flows.",
					3));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"XNPV",
					"XNPV(rate, values, dates)",
					"Returns the net present value for a schedule of case flows.",
					3));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"YIELD",
					"YIELD(settlement, maturity, rate, pr, redemption, frequency, basus)",
					"Returns the yield on a security that pays periodic interest.",
					7));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"YIELDDISC",
					"YIELDDISC(settlement, maturity, pr, redemption, basis)",
					"Returns the annual yield for a discounted security. For example, a treasury bill.",
					5));
		
		formulaInfos.put(FINANCIAL, 
				new FormulaMetaInfo(FINANCIAL,
					"YIELDMAT",
					"YIELDMAT(settlement, maturity, issue, rate, pr, basis)",
					"Returns the annual yield of a security that pays interest at maturity",
					6));
		
	}
}
