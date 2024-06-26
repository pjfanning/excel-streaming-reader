package com.github.pjfanning.xlsx.exceptions;

public class OpenException extends ExcelRuntimeException {

  public OpenException() {
    super();
  }

  public OpenException(String msg) {
    super(msg);
  }

  public OpenException(Exception e) {
    super(e);
  }

  public OpenException(String msg, Exception e) {
    super(msg, e);
  }
}
