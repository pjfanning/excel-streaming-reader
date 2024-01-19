package com.github.pjfanning.xlsx.exceptions;

/**
 * A parent class for all the excel-streaming-reader specific Runtime Exceptions.
 *
 * @since 4.3.0
 */
public class ExcelRuntimeException extends RuntimeException {

  protected ExcelRuntimeException() {
    super();
  }

  protected ExcelRuntimeException(String msg) {
    super(msg);
  }

  protected ExcelRuntimeException(Exception e) {
    super(e);
  }

  protected ExcelRuntimeException(String msg, Exception e) {
    super(msg, e);
  }
}
