package org.zkoss.zss.engine;

import org.zkoss.lang.Objects;

public class Complex extends org.apache.commons.math.complex.Complex {
	
	protected String suffix;
	
	public Complex(double real, double imaginary, String suffix) {
		super(real, imaginary);
		this.suffix = suffix;
	}
	
	public Complex(double real, double imaginary){
		this(real, imaginary, "i");
	}

	public int hashCode() {
		return (suffix == null ? 0 : suffix.hashCode()) ^ super.hashCode();
	}
	public boolean equals(Object arg0) {
		return (this == arg0)
			|| (arg0 instanceof Complex 
					&& super.equals(arg0) 
					&& Objects.equals(this.suffix, ((Complex)arg0).getSuffix()));
	}

	public String getSuffix() {
		return suffix;
	}
}
