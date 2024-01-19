package com.github.pjfanning.xlsx.exceptions;

public class CloseException extends ExcelRuntimeException {

  public CloseException() {
    super();
  }

  public CloseException(String msg) {
    super(msg);
  }

  public CloseException(Exception e) {
    super(e);
  }

  public CloseException(String msg, Exception e) {
    super(msg, e);
  }
}
