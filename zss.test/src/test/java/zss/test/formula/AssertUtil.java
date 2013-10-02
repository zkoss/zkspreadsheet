package zss.test.formula;

import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.apache.commons.math.complex.Complex;
import org.junit.Assert;
import org.zkoss.poi.ss.formula.functions.ComplexFormat;
import org.zkoss.zss.model.Range;

public class AssertUtil {

	public static double DELTA = 1E-8; // formula precision setting : 10 ^ -8

	/**
	 * test whether two complex is equals 
	 */
	public static void assertComplexEquals(String complex1, String complex2) {

		try {

			complex1 = replaceiWith1i(complex1);
			complex2 = replaceiWith1i(complex2);

			ComplexFormat cf = new ComplexFormat();
			Complex c1 = cf.parse(complex1);
			Complex c2 = cf.parse(complex2);

			double real1 = c1.getReal();
			double im1 = c1.getImaginary();
			double real2 = c2.getReal();
			double im2 = c2.getImaginary();

			if(Math.abs(real1 - real2) <= DELTA && Math.abs(im1 - im2) <= DELTA) {
				return;
			}

		} catch (ParseException e) {
			e.printStackTrace();
		}

		org.junit.Assert.fail(complex1 + " is not equals to " + complex2);;

	}

	/**
	 * only accept i, j, k as imaginary character
	 */
	private static String getComplexImaginaryChar(String complex) {
		String s = "i";
		int i = complex.indexOf(s);
		if (i == -1) {
			s = "j";
			i = complex.indexOf(s);
			if (i == -1) {
				s = "k";
				i = complex.indexOf(s);
				if (i == -1) {
					throw new RuntimeException("only accept i, j, k as imaginary character");
				}
				return s; // k
			}
			return s; // j
		}
		return s; // i
	}

	/**
	 * Apache common cannot parse complex like "8+i", must replace i as 1i.
	 * So the result string will become "8+1i".
	 * Only replace when the (indexOf "i" - 1) is empty or operator.
	 * There are two example case:
	 * 8+i, i.  
	 */
	private static String replaceiWith1i(String complex) {

		String imChar = getComplexImaginaryChar(complex);

		if(complex.indexOf(imChar) == 0) {
			return complex.replace(imChar, "1" + imChar);
		}

		if(complex.charAt(complex.indexOf(imChar) - 1) == '+' || complex.charAt(complex.indexOf(imChar) - 1) == '-') {
			return complex.replace(imChar, "1" + imChar);
		}

		return complex;
	}

}