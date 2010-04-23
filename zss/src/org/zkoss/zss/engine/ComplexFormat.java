package org.zkoss.zss.engine;

import java.text.NumberFormat;
import java.text.ParseException;

import org.apache.commons.math.complex.Complex;

public class ComplexFormat extends
		org.apache.commons.math.complex.ComplexFormat {

	public ComplexFormat() {
		super();
		// TODO Auto-generated constructor stub
	}

	public ComplexFormat(NumberFormat realFormat, NumberFormat imaginaryFormat) {
		super(realFormat, imaginaryFormat);
		// TODO Auto-generated constructor stub
	}

	public ComplexFormat(NumberFormat format) {
		super(format);
		// TODO Auto-generated constructor stub
	}

	public ComplexFormat(String imaginaryCharacter, NumberFormat realFormat,
			NumberFormat imaginaryFormat) {
		super(imaginaryCharacter, realFormat, imaginaryFormat);
		// TODO Auto-generated constructor stub
	}

	public ComplexFormat(String imaginaryCharacter, NumberFormat format) {
		super(imaginaryCharacter, format);
		// TODO Auto-generated constructor stub
	}

	public ComplexFormat(String imaginaryCharacter) {
		super(imaginaryCharacter);
		// TODO Auto-generated constructor stub
	}

	public org.zkoss.zss.engine.Complex parse(String source, String suffix) throws ParseException {
		Complex c = super.parse(source);
		org.zkoss.zss.engine.Complex result = new org.zkoss.zss.engine.Complex(c.getReal(), c.getImaginary(), suffix);
		return result;
	}

}
