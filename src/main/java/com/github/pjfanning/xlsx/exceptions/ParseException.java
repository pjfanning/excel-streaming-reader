package com.github.pjfanning.xlsx.exceptions;

public class ParseException extends ExcelRuntimeException {

  public ParseException() {
    super();
  }

  public ParseException(String msg) {
    super(msg);
  }

  public ParseException(Exception e) {
    super(e);
  }

  public ParseException(String msg, Exception e) {
    super(msg, e);
  }
}
