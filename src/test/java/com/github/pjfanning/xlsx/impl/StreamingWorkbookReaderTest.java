package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.StreamingReader;
import org.junit.Test;

public class StreamingWorkbookReaderTest {

  /**
   * close() used to throw NullPointerException if the reader was never successfully initialised,
   * because the OPCPackage field was still null.
   */
  @Test
  public void testCloseBeforeInit() throws Exception {
    StreamingWorkbookReader reader = new StreamingWorkbookReader(StreamingReader.builder());
    reader.close();
    //closing twice should also be safe
    reader.close();
  }
}
