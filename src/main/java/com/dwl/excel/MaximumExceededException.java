package com.dwl.excel;

public class MaximumExceededException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  public MaximumExceededException(String message) {
    super(message);
  }

  public MaximumExceededException(String message, Throwable cause) {
    super(message, cause);
  }
}
