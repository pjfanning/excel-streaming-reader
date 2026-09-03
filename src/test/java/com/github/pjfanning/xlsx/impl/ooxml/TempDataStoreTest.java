package com.github.pjfanning.xlsx.impl.ooxml;

import org.junit.Test;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Random;

import static org.junit.Assert.assertArrayEquals;

public class TempDataStoreTest {

  /**
   * The output stream is buffered, so anything still sitting in the buffer has to reach the file
   * before getInputStream() is called. A payload smaller than the buffer would be lost entirely
   * if that ever stopped being true.
   */
  @Test
  public void testTempFileDataStoreRoundTrip() throws Exception {
    assertRoundTrip(new TempFileDataStore(), 64);
    assertRoundTrip(new TempFileDataStore(), 8192 * 3 + 17);
  }

  @Test
  public void testTempMemoryDataStoreRoundTrip() throws Exception {
    assertRoundTrip(new TempMemoryDataStore(), 64);
    assertRoundTrip(new TempMemoryDataStore(), 8192 * 3 + 17);
  }

  private void assertRoundTrip(TempDataStore store, int size) throws Exception {
    final byte[] payload = new byte[size];
    new Random(size).nextBytes(payload);
    try {
      try (OutputStream os = store.getOutputStream()) {
        os.write(payload);
      }
      try (InputStream is = store.getInputStream();
           ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
        final byte[] buf = new byte[1024];
        int read;
        while ((read = is.read(buf)) != -1) {
          bos.write(buf, 0, read);
        }
        assertArrayEquals(payload, bos.toByteArray());
      }
    } finally {
      store.close();
    }
  }
}
