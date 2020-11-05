package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.poi.xssf.streaming.TempFileSharedStringsTable;
import com.github.pjfanning.xlsx.StreamingReader.Builder;
import com.github.pjfanning.xlsx.XmlUtils;
import com.github.pjfanning.xlsx.exceptions.NotSupportedException;
import com.github.pjfanning.xlsx.exceptions.OpenException;
import com.github.pjfanning.xlsx.exceptions.ParseException;
import com.github.pjfanning.xlsx.exceptions.ReadException;
import com.github.pjfanning.xlsx.impl.ooxml.OoXmlStrictConverterInputStream;
import java.io.FileInputStream;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static com.github.pjfanning.xlsx.XmlUtils.searchForNodeList;
import static java.util.Arrays.asList;

public class StreamingWorkbookReader implements Iterable<Sheet>, AutoCloseable {
  private static final Logger log = LoggerFactory.getLogger(StreamingWorkbookReader.class);

  private final List<StreamingSheet> sheets;
  private final List<Map<String, String>> sheetProperties = new ArrayList<>();
  private final Builder builder;
  private File tmp;
  private OPCPackage pkg;
  private SharedStringsTable sst;
  private boolean use1904Dates = false;
  private StreamingWorkbook workbook = null;
  private POIXMLProperties.CoreProperties coreProperties = null;

  /**
   * This constructor exists only so the StreamingReader can instantiate
   * a StreamingWorkbook using its own reader implementation. Do not use
   * going forward.
   *
   * @param sst      The SST data for this workbook
   * @param pkg      The POI package that should be closed when this workbook is closed
   * @param reader   A single streaming reader instance
   * @param builder  The builder containing all options
   */
  @Deprecated
  public StreamingWorkbookReader(SharedStringsTable sst, OPCPackage pkg, StreamingSheetReader reader, Builder builder) {
    this.sst = sst;
    this.pkg = pkg;
    this.sheets = asList(new StreamingSheet(null, null, reader));
    this.builder = builder;
  }

  public StreamingWorkbookReader(Builder builder) {
    this.sheets = new ArrayList<>();
    this.builder = builder;
  }

  public StreamingSheetReader first() {
    return sheets.get(0).getReader();
  }

  public void init(InputStream is) {
    File f = null;
    try {
      f = TempFileUtil.writeInputStreamToFile(is, builder.getBufferSize());
      log.debug("Created temp file [" + f.getAbsolutePath() + "]");

      init(f);
      tmp = f;
    } catch(IOException e) {
      throw new ReadException("Unable to read input stream", e);
    } catch(RuntimeException e) {
      if (f != null) {
        f.delete();
      }
      throw e;
    }
  }

  public void init(File f) {
    try {
      if(builder.getPassword() != null) {
        // Based on: https://poi.apache.org/encryption.html
        POIFSFileSystem poifs = new POIFSFileSystem(f);
        EncryptionInfo info = new EncryptionInfo(poifs);
        Decryptor d = Decryptor.getInstance(info);
        d.verifyPassword(builder.getPassword());
        if (builder.convertFromOoXmlStrict()) {
          pkg = OPCPackage.open(new OoXmlStrictConverterInputStream(d.getDataStream(poifs)));
        } else {
          pkg = OPCPackage.open(d.getDataStream(poifs));
        }
      } else {
        if (builder.convertFromOoXmlStrict()) {
          pkg = OPCPackage.open(new OoXmlStrictConverterInputStream(new FileInputStream(f)));
        } else {
          pkg = OPCPackage.open(f);
        }
      }

      XSSFReader reader = new XSSFReader(pkg);
      if(builder.useSstTempFile()) {
        log.debug("Created sst cache file");
        sst = new TempFileSharedStringsTable(pkg, builder.encryptSstTempFile());
      } else {
        sst = reader.getSharedStringsTable();
      }

      if (builder.readCoreProperties()) {
        try {
          POIXMLProperties xmlProperties = new POIXMLProperties(pkg);
          coreProperties = xmlProperties.getCoreProperties();
        } catch (Exception e) {
          log.warn("Failed to read coreProperties", e);
        }
      }

      StylesTable styles = reader.getStylesTable();
      use1904Dates = WorkbookUtil.use1904Dates(reader);

      loadSheets(reader, sst, styles, builder.getRowCacheSize());
    } catch(SAXException | ParserConfigurationException e) {
      throw new ParseException("Failed to parse file", e);
    } catch(IOException e) {
      throw new OpenException("Failed to open file", e);
    } catch(OpenXML4JException | XMLStreamException e) {
      throw new ReadException("Unable to read workbook", e);
    } catch(GeneralSecurityException e) {
      throw new ReadException("Unable to read workbook - Decryption failed", e);
    }
  }

  void setWorkbook(StreamingWorkbook workbook) {
    this.workbook = workbook;
    workbook.setCoreProperties(coreProperties);
  }

  void loadSheets(XSSFReader reader, SharedStringsTable sst, StylesTable stylesTable, int rowCacheSize) throws IOException, InvalidFormatException,
      XMLStreamException {
    lookupSheetNames(reader);

    //Some workbooks have multiple references to the same sheet. Need to filter
    //them out before creating the XMLEventReader by keeping track of their URIs.
    //The sheets are listed in order, so we must keep track of insertion order.
    SheetIterator iter = (SheetIterator) reader.getSheetsData();
    Map<URI, InputStream> sheetStreams = new LinkedHashMap<>();
    while(iter.hasNext()) {
      InputStream is = iter.next();
      sheetStreams.put(iter.getSheetPart().getPartName().getURI(), is);
    }

    //Iterate over the loaded streams
    int i = 0;
    for(URI uri : sheetStreams.keySet()) {
      XMLEventReader parser = XMLHelper.newXMLInputFactory().createXMLEventReader(sheetStreams.get(uri));
      sheets.add(new StreamingSheet(
              workbook,
              sheetProperties.get(i++).get("name"),
              new StreamingSheetReader(sst, stylesTable, parser, use1904Dates, rowCacheSize)));
    }
  }

  void lookupSheetNames(XSSFReader reader) throws IOException, InvalidFormatException {
    sheetProperties.clear();
    try {
      NodeList nl = searchForNodeList(XmlUtils.readDocument(reader.getWorkbookData()), "/ss:workbook/ss:sheets/ss:sheet");
      for(int i = 0; i < nl.getLength(); i++) {
        Map<String, String> props = new HashMap<>();
        props.put("name", nl.item(i).getAttributes().getNamedItem("name").getTextContent());

        Node state = nl.item(i).getAttributes().getNamedItem("state");
        props.put("state", state == null ? "visible" : state.getTextContent());
        sheetProperties.add(props);
      }
    } catch (SAXException|ParserConfigurationException e) {
      throw new ParseException("Failed to parse file", e);
    }
  }

  List<? extends Sheet> getSheets() {
    return sheets;
  }

  public List<Map<String, String>> getSheetProperties() {
    return sheetProperties;
  }

  @Override
  public Iterator<Sheet> iterator() {
    return new StreamingSheetIterator(sheets.iterator());
  }

  @Override
  public void close() throws IOException {
    try {
      for(StreamingSheet sheet : sheets) {
        sheet.getReader().close();
      }
      pkg.revert();
    } finally {
      if(tmp != null) {
        if (log.isDebugEnabled()) {
          log.debug("Deleting tmp file [" + tmp.getAbsolutePath() + "]");
        }
        tmp.delete();
      }
      if(sst != null) {
        sst.close();
      }
    }
  }

  static class StreamingSheetIterator implements Iterator<Sheet> {
    private final Iterator<StreamingSheet> iterator;

    public StreamingSheetIterator(Iterator<StreamingSheet> iterator) {
      this.iterator = iterator;
    }

    @Override
    public boolean hasNext() {
      return iterator.hasNext();
    }

    @Override
    public Sheet next() {
      return iterator.next();
    }

    @Override
    public void remove() {
      throw new NotSupportedException();
    }
  }
}
