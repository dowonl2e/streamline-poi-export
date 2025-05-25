package com.dwl.excel;

public class NotAllowedClassException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  public NotAllowedClassException(String message) {
    super(message);
  }

  public NotAllowedClassException(String message, Throwable cause) {
    super(message, cause);
  }

}
