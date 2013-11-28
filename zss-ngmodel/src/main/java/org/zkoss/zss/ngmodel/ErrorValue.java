package org.zkoss.zss.ngmodel;

public class ErrorValue {

	public static final byte INVALID_FORMULA = (byte)0;
	public static final byte INVALID_VALUE = (byte)1;
	public static final byte INVALID_NAME = (byte)2;
	
	
	private byte code;
	private String message;

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
	
	public String gettErrorString(){
		if(code==INVALID_NAME){
			return "#NAME!";
		}else if(code==INVALID_VALUE){
			return "#VALUE!";
		}
		return "#NAME!";
	}

}
