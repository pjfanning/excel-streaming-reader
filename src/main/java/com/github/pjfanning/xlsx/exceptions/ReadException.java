package com.github.pjfanning.xlsx.exceptions;

public class ReadException extends ExcelRuntimeException {

  public ReadException() {
    super();
  }

  public ReadException(String msg) {
    super(msg);
  }

  public ReadException(Exception e) {
    super(e);
  }

  public ReadException(String msg, Exception e) {
    super(msg, e);
  }
}
