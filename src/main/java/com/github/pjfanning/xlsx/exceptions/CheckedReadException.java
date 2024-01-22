package com.github.pjfanning.xlsx.exceptions;

/**
 * An exception that is thrown if there is a problem reading the Excel file.
 *
 * @see ReadException
 * @since 4.3.0
 */
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
