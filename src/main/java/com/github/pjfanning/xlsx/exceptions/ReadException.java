package com.github.pjfanning.xlsx.exceptions;

/**
 * A Runtime Exception that is thrown if there is a problem reading the Excel file.
 * This is used in APIs where the method is unable to throw a checked exception.
 *
 * @see CheckedReadException
 */
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
