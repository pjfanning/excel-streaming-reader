package com.github.pjfanning.xlsx.exceptions;

/**
 * A checked exception that is thrown if there is a problem reading the Excel file.
 *
 * <p>
 *   To avoid adding a large number of new checked exceptions to the method signatures,
 *   this exception is generic. Any read issue will throw this exception. You can look call
 *   {@link #getCause()} to get the underlying exception will is likely to be one of the more specific
 *   legacy exceptions that implement {@link com.github.pjfanning.xlsx.exceptions.ExcelRuntimeException}.
 * </p>
 *
 * @see ReadException
 * @since 4.4.0
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
