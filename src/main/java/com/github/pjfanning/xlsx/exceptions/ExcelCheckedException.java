package com.github.pjfanning.xlsx.exceptions;

/**
 * A parent class for all the excel-streaming-reader specific Checked Exceptions (i.e. not Runtime Exceptions).
 *
 * @since 4.3.0
 */
public class ExcelCheckedException extends Exception {

    protected ExcelCheckedException() {
        super();
    }

    protected ExcelCheckedException(String msg) {
        super(msg);
    }

    protected ExcelCheckedException(Exception e) {
        super(e);
    }

    protected ExcelCheckedException(String msg, Exception e) {
        super(msg, e);
    }
}
