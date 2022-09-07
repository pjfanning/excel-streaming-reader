package com.github.pjfanning.xlsx.impl;

import org.apache.poi.util.TempFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public final class TempFileUtil {
  private static final Logger log = LoggerFactory.getLogger(TempFileUtil.class);

  private TempFileUtil() {}

  public static File writeInputStreamToFile(InputStream is, int bufferSize) throws IOException {
    if (is == null) throw new NullPointerException("InputStream is null");
    File f = TempFile.createTempFile("tmp-", ".xlsx");
    try (FileOutputStream fos = new FileOutputStream(f)) {
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
