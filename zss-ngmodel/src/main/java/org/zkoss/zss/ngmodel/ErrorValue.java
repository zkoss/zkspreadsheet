package org.zkoss.zss.ngmodel;

public class ErrorValue {

	byte code;
	String message;

	public ErrorValue(byte code) {
		this(code, null);
	}

	public ErrorValue(byte code, String message) {
		this.code = code;
		this.message = message;
	}

	public byte getCode() {
		return code;
	}

	public void setCode(byte code) {
		this.code = code;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String message) {
		this.message = message;
	}

}
