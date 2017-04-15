package com.exceptions;

public class XLSFatalException extends Exception{
	
	public XLSFatalException() {
		super();
	}
	public XLSFatalException(String message) {
		super(message);
	}
	
	public XLSFatalException(Throwable cause) {
        super(cause);
    }

    public XLSFatalException(String message, Throwable cause) {
        super(message, cause);
    } 
	
	public String getMessage() {
		return super.getMessage();
	}
}
