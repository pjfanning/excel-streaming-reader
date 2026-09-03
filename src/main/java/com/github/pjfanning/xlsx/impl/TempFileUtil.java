package com.github.pjfanning.xlsx.impl;

import org.apache.poi.util.TempFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

public class TempFileUtil {
  private static final Logger log = LoggerFactory.getLogger(TempFileUtil.class);
  //BufferedOutputStream passes writes of at least this size straight through, so a caller that
  //asks for a large bufferSize does not pay for an extra copy
  private static final int WRITE_BUFFER_SIZE = 8192;

  private TempFileUtil() {}

  public static File writeInputStreamToFile(InputStream is, int bufferSize) throws IOException {
    if (is == null) throw new NullPointerException("InputStream is null");
    File f = TempFile.createTempFile("tmp-", ".xlsx");
    try (OutputStream fos = new BufferedOutputStream(new FileOutputStream(f), WRITE_BUFFER_SIZE)) {
      int read;
      byte[] bytes = new byte[bufferSize];
      while ((read = is.read(bytes)) != -1) {
        fos.write(bytes, 0, read);
      }
      return f;
    } catch (IOException | RuntimeException | Error e) {
      try {
        if(!f.delete()) {
          log.debug("failed to delete temp file");
        }
      } catch (Exception fileException) {
        log.warn("Failed to delete temp file {}: {}", f.getAbsolutePath(), fileException.toString());
      }
      throw e;
    } finally {
      is.close();
    }
  }
}
