package com.github.pjfanning.xlsx.exceptions;

public class CheckedReadException extends ExcelCheckedException {

  public CheckedReadException() {
    super();
  }

  public CheckedReadException(String msg) {
    super(msg);
  }

  public CheckedReadException(Exception e) {
    super(e);
  }

  public CheckedReadException(String msg, Exception e) {
    super(msg, e);
  }
}
