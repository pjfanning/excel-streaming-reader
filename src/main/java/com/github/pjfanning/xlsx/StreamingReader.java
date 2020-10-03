package com.github.pjfanning.xlsx;

import com.github.pjfanning.poi.xssf.streaming.TempFileSharedStringsTable;
import com.github.pjfanning.xlsx.exceptions.*;
import com.github.pjfanning.xlsx.impl.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.Iterator;
import java.util.Objects;

import static com.github.pjfanning.xlsx.XmlUtils.searchForNodeList;

/**
 * Streaming Excel workbook implementation. Most advanced features of POI are not supported.
 * Use this only if your application can handle iterating through an entire workbook, row by
 * row.
 */
public class StreamingReader implements Iterable<Row>, AutoCloseable {
  private static final Logger log = LoggerFactory.getLogger(StreamingReader.class);

  private File tmp;
  private final StreamingWorkbookReader workbook;

  public StreamingReader(StreamingWorkbookReader workbook) {
    this.workbook = workbook;
  }

  /**
   * Returns a new streaming iterator to loop through rows. This iterator is not
   * guaranteed to have all rows in memory, and any particular iteration may
   * trigger a load from disk to read in new data.
   *
   * @return the streaming iterator
   * @deprecated StreamingReader is equivalent to the POI Workbook object rather
   * than the Sheet object. This method will be removed in a future release.
   */
  @Override
  public Iterator<Row> iterator() {
    return workbook.first().iterator();
  }

  /**
   * Closes the streaming resource, attempting to clean up any temporary files created.
   *
   * @throws CloseException if there is an issue closing the stream
   */
  @Override
  public void close() throws IOException {
    try {
      workbook.close();
    } finally {
      if(tmp != null) {
        if (log.isDebugEnabled()) {
          log.debug("Deleting tmp file [" + tmp.getAbsolutePath() + "]");
        }
        tmp.delete();
      }
    }
  }

  public static Builder builder() {
    return new Builder();
  }

  public static class Builder {
    private int rowCacheSize = 10;
    private int bufferSize = 1024;
    private int sheetIndex = 0;
    private boolean useSstTempFile = false;
    private boolean encryptSstTempFile = false;
    private boolean readCoreProperties = false;
    private String sheetName;
    private String password;

    public int getRowCacheSize() {
      return rowCacheSize;
    }

    public int getBufferSize() {
      return bufferSize;
    }

    /**
     * @return The sheet index
     * @deprecated This method will be removed in a future release.
     */
    public int getSheetIndex() {
      return sheetIndex;
    }

    /**
     * @return The sheet name
     * @deprecated This method will be removed in a future release.
     */
    public String getSheetName() {
      return sheetName;
    }

    /**
     * @return The password to use to unlock this workbook
     */
    public String getPassword() {
      return password;
    }

    /**
     * @return Whether to use a temp file for the Shared Strings data. If false, no
     * temp file will be used and the entire table will be loaded into memory.
     */
    public boolean useSstTempFile() {
      return useSstTempFile;
    }

    /**
     * @return Whether to encrypt the temp file for the Shared Strings data. Only applies if <code>useSstTempFile()</code>
     * is true.
     */
    public boolean encryptSstTempFile() {
      return encryptSstTempFile;
    }

    /**
     * @return Whether to read the core document properties.
     */
    public boolean readCoreProperties() {
      return readCoreProperties;
    }

    /**
     * The number of rows to keep in memory at any given point.
     * <p>
     * Defaults to 10
     * </p>
     *
     * @param rowCacheSize number of rows
     * @return reference to current {@code Builder}
     */
    public Builder rowCacheSize(int rowCacheSize) {
      this.rowCacheSize = rowCacheSize;
      return this;
    }

    /**
     * The number of bytes to read into memory from the input
     * resource.
     * <p>
     * Defaults to 1024
     * </p>
     *
     * @param bufferSize buffer size in bytes
     * @return reference to current {@code Builder}
     */
    public Builder bufferSize(int bufferSize) {
      this.bufferSize = bufferSize;
      return this;
    }

    /**
     * Which sheet to open. There can only be one sheet open
     * for a single instance of {@code StreamingReader}. If
     * more sheets need to be read, a new instance must be
     * created.
     * <p>
     * Defaults to 0
     * </p>
     *
     * @param sheetIndex index of sheet
     * @return reference to current {@code Builder}
     * @deprecated This method will be removed in a future release. Use {@link StreamingWorkbook#getSheetAt(int)} instead.
     */
    public Builder sheetIndex(int sheetIndex) {
      this.sheetIndex = sheetIndex;
      return this;
    }

    /**
     * Which sheet to open. There can only be one sheet open
     * for a single instance of {@code StreamingReader}. If
     * more sheets need to be read, a new instance must be
     * created.
     *
     * @param sheetName name of sheet
     * @return reference to current {@code Builder}
     * @deprecated This method will be removed in a future release. Use {@link StreamingWorkbook#getSheet(String)} instead.
     */
    public Builder sheetName(String sheetName) {
      this.sheetName = sheetName;
      return this;
    }

    /**
     * For password protected files specify password to open file.
     * If the password is incorrect a {@code ReadException} is thrown on
     * {@code read}.
     * <p>NULL indicates that no password should be used, this is the
     * default value.</p>
     *
     * @param password to use when opening file
     * @return reference to current {@code Builder}
     */
    public Builder password(String password) {
      this.password = password;
      return this;
    }

    /**
     * Enables use of Shared Strings Table temp file. This option exists to accommodate
     * extremely large workbooks with millions of unique strings. Normally the SST is entirely
     * loaded into memory, but with large workbooks with high cardinality (i.e., very few
     * duplicate values) the SST may not fit entirely into memory.
     * <p>
     * By default, the entire SST *will* be loaded into memory. <strong>However</strong>,
     * enabling this option at all will have some noticeable performance degradation as you are
     * trading memory for disk space.
     *
     * @param useSstTempFile whether to use a temp file to store the Shared Strings Table data
     * @return reference to current {@code Builder}
     */
    public Builder setUseSstTempFile(boolean useSstTempFile) {
      this.useSstTempFile = useSstTempFile;
      return this;
    }

    /**
     * Enables use of encryption in Shared Strings Table temp file. This only applies if <code>setUseSstTempFile</code>
     * is set to true.
     * <p>
     * By default, the temp file is not encrypted. <strong>However</strong>,
     * enabling this option could slow down the processing of Shared Strings data.
     *
     * @param encryptSstTempFile whether to encrypt the temp file used to store the Shared Strings Table data
     * @return reference to current {@code Builder}
     */
    public Builder setEncryptSstTempFile(boolean encryptSstTempFile) {
      this.encryptSstTempFile = encryptSstTempFile;
      return this;
    }

    /**
     * Enables the reading of the core document properties.
     *
     * @param readCoreProperties whether to read the core document properties
     * @return reference to current {@code Builder}
     */
    public Builder setReadCoreProperties(boolean readCoreProperties) {
      this.readCoreProperties = readCoreProperties;
      return this;
    }

    /**
     * Reads a given {@code InputStream} and returns a new
     * instance of {@code Workbook}. Due to Apache POI
     * limitations, a temporary file must be written in order
     * to create a streaming iterator. This process will use
     * the same buffer size as specified in {@link #bufferSize(int)}.
     *
     * @param is input stream to read in
     * @return A {@link Workbook} that can be read from
     * @throws ReadException if there is an issue reading the stream
     */
    public Workbook open(InputStream is) {
      StreamingWorkbookReader workbook = new StreamingWorkbookReader(this);
      workbook.init(is);
      return new StreamingWorkbook(workbook);
    }

    /**
     * Reads a given {@code File} and returns a new instance
     * of {@code Workbook}.
     *
     * @param file file to read in
     * @return built streaming reader instance
     * @throws OpenException if there is an issue opening the file
     * @throws ReadException if there is an issue reading the file
     */
    public Workbook open(File file) {
      StreamingWorkbookReader workbook = new StreamingWorkbookReader(this);
      workbook.init(file);
      return new StreamingWorkbook(workbook);
    }

    /**
     * Reads a given {@code InputStream} and returns a new
     * instance of {@code StreamingReader}. Due to Apache POI
     * limitations, a temporary file must be written in order
     * to create a streaming iterator. This process will use
     * the same buffer size as specified in {@link #bufferSize(int)}.
     *
     * @param is input stream to read in
     * @return built streaming reader instance
     * @throws ReadException if there is an issue reading the stream
     * @deprecated This method will be removed in a future release. Use {@link Builder#open(InputStream)} instead
     */
    public StreamingReader read(InputStream is) {
      File f = null;
      try {
        f = TempFileUtil.writeInputStreamToFile(is, bufferSize);
        log.debug("Created temp file [" + f.getAbsolutePath() + "]");

        StreamingReader r = read(f);
        r.tmp = f;
        return r;
      } catch(IOException e) {
        throw new ReadException("Unable to read input stream", e);
      } catch(RuntimeException e) {
        if (f != null) {
          f.delete();
        }
        throw e;
      }
    }

    /**
     * Reads a given {@code File} and returns a new instance
     * of {@code StreamingReader}.
     *
     * @param f file to read in
     * @return built streaming reader instance
     * @throws OpenException if there is an issue opening the file
     * @throws ReadException if there is an issue reading the file
     * @deprecated This method will be removed in a future release. Use {@link Builder#open(File)} instead
     */
    public StreamingReader read(File f) {
      try {
        OPCPackage pkg;
        if(password != null) {
          // Based on: https://poi.apache.org/encryption.html
          POIFSFileSystem poifs = new POIFSFileSystem(f);
          EncryptionInfo info = new EncryptionInfo(poifs);
          Decryptor d = Decryptor.getInstance(info);
          d.verifyPassword(password);
          pkg = OPCPackage.open(d.getDataStream(poifs));
        } else {
          pkg = OPCPackage.open(f);
        }

        boolean use1904Dates = false;
        XSSFReader reader = new XSSFReader(pkg);

        SharedStringsTable sst;
        if(useSstTempFile) {
          log.debug("Created sst cache file");
          sst = new TempFileSharedStringsTable(pkg, encryptSstTempFile);
        } else {
          sst = reader.getSharedStringsTable();
        }

        StylesTable styles = reader.getStylesTable();
        use1904Dates = WorkbookUtil.use1904Dates(reader);
        InputStream sheet = findSheet(reader);
        if(sheet == null) {
          throw new MissingSheetException("Unable to find sheet at index [" + sheetIndex + "]");
        }

        XMLEventReader parser = XMLHelper.newXMLInputFactory().createXMLEventReader(sheet);

        return new StreamingReader(new StreamingWorkbookReader(sst, pkg, new StreamingSheetReader(sst, styles, parser, use1904Dates, rowCacheSize),
            this));
      } catch(SAXException|ParserConfigurationException e) {
        throw new ParseException("Failed to parse file", e);
      } catch(IOException e) {
        throw new OpenException("Failed to open file", e);
      } catch(OpenXML4JException | XMLStreamException e) {
        throw new ReadException("Unable to read workbook", e);
      } catch(GeneralSecurityException e) {
        throw new ReadException("Unable to read workbook - Decryption failed", e);
      }
    }

    /**
     * @deprecated This will be removed when the transition to the 1.x API is complete
     */
    private InputStream findSheet(XSSFReader reader) throws IOException, InvalidFormatException {
      int index = sheetIndex;
      if(sheetName != null) {
        index = -1;
        try {
          //This file is separate from the worksheet data, and should be fairly small
          NodeList nl = searchForNodeList(XmlUtils.readDocument(reader.getWorkbookData()), "/ss:workbook/ss:sheets/ss:sheet");
          for (int i = 0; i < nl.getLength(); i++) {
            if (Objects.equals(nl.item(i).getAttributes().getNamedItem("name").getTextContent(), sheetName)) {
              index = i;
            }
          }
        } catch (SAXException| ParserConfigurationException e) {
          throw new ParseException("Failed to parse file", e);
        }
        if(index < 0) {
          return null;
        }
      }
      Iterator<InputStream> iter = reader.getSheetsData();
      InputStream sheet = null;

      int i = 0;
      while(iter.hasNext()) {
        InputStream is = iter.next();
        if(i++ == index) {
          sheet = is;
          log.debug("Found sheet at index [" + sheetIndex + "]");
          break;
        }
      }
      return sheet;
    }
  }

}
