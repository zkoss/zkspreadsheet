package org.zkoss.zss.engine;

public class Complex extends org.apache.commons.math.complex.Complex {
	
	protected String suffix;
	
	public Complex(double real, double imaginary, String suffix) {
		super(real, imaginary);
		this.suffix = suffix;
	}
	
	public Complex(double real, double imaginary){
		this(real, imaginary, "i");
	}

	public boolean equals(Object arg0) {
		boolean ret = super.equals(arg0);
		if(ret) {
			Complex rhs = (Complex)arg0;
			if(!this.suffix.equals(rhs.getSuffix())) {
				ret = false;
			}			
		}
		return ret;
	}

	public String getSuffix() {
		return suffix;
	}
}
