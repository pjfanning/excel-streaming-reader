package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.exceptions.MissingSheetException;
import com.github.pjfanning.xlsx.exceptions.ReadException;
import com.github.pjfanning.xlsx.impl.adapter.WorkbookAdapter;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

import javax.xml.stream.XMLStreamException;
import java.io.IOException;
import java.util.*;
import java.util.regex.Pattern;

public class StreamingWorkbook implements WorkbookAdapter, Date1904Support, AutoCloseable {
  private final StreamingWorkbookReader reader;
  private POIXMLProperties.CoreProperties coreProperties = null;
  private List<XSSFPictureData> pictures;

  public StreamingWorkbook(StreamingWorkbookReader reader) {
    this.reader = reader;
    reader.setWorkbook(this);
  }

  int findSheetByName(final String name) {
    final List<Map<String, String>> props = reader.getSheetProperties();
    final int size = props.size();
    for(int i = 0; i < size; i++) {
      if(name.equalsIgnoreCase(props.get(i).get("name"))) {
        return i;
      }
    }
    return -1;
  }

  /* Supported */

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Sheet> iterator() {
    return reader.iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Iterator<Sheet> sheetIterator() {
    return iterator();
  }

  @Override
  public Spliterator<Sheet> spliterator() {
    return reader.spliterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public String getSheetName(int sheet) {
    return reader.getSheetProperties().get(sheet).get("name");
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getSheetIndex(String name) {
    return findSheetByName(name);
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getSheetIndex(Sheet sheet) {
    if(sheet instanceof StreamingSheet) {
      return findSheetByName(sheet.getSheetName());
    } else {
      throw new UnsupportedOperationException("Cannot use non-StreamingSheet sheets");
    }
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public int getNumberOfSheets() {
    return reader.getSheetProperties().size();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Sheet getSheetAt(final int index) {
    try {
      return reader.getSheetAt(index);
    } catch (XMLStreamException|IOException e) {
      throw new ReadException(e);
    }
  }

  /**
   * Get sheet with the given name
   *
   * @param name - of the sheet
   * @return Sheet with the name provided
   * @throws MissingSheetException if no sheet is found with the provided <code>name</code>
   */
  @Override
  public Sheet getSheet(String name) {
    try {
      return reader.getSheet(name);
    } catch (XMLStreamException|IOException e) {
      throw new ReadException(e);
    } catch (NoSuchElementException nse) {
      throw new MissingSheetException("Failed to find sheet: " + name);
    }
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public boolean isSheetHidden(int sheetIx) {
    return "hidden".equals(reader.getSheetProperties().get(sheetIx).get("state"));
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public boolean isSheetVeryHidden(int sheetIx) {
    return "veryHidden".equals(reader.getSheetProperties().get(sheetIx).get("state"));
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public SpreadsheetVersion getSpreadsheetVersion() {
    return SpreadsheetVersion.EXCEL2007;
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public void close() throws IOException {
    reader.close();
  }

  /**
   * Returns the Core Properties if this feature is enabled on the {@link com.github.pjfanning.xlsx.StreamingReader.Builder}
   *
   * @return CoreProperties
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadCoreProperties(boolean)} is not set to true
   */
  public POIXMLProperties.CoreProperties getCoreProperties() {
    if (reader.getBuilder().readCoreProperties()) {
      return coreProperties;
    } else {
      throw new IllegalStateException("getCoreProperties() only works if StreamingWorking.Builder setReadCoreProperties is set to true");
    }
  }

  void setCoreProperties(POIXMLProperties.CoreProperties coreProperties) {
    this.coreProperties = coreProperties;
  }

  /**
   * Gets all pictures from the Workbook. This approach is not stream friendly.
   *
   * @return the list of pictures (a list of {@link XSSFPictureData} objects.)
   */
  @Override
  public List<? extends PictureData> getAllPictures() {
    if(pictures == null){
      List<PackagePart> mediaParts = reader.getOPCPackage().getPartsByName(Pattern.compile("/xl/media/.*?"));
      pictures = new ArrayList<>(mediaParts.size());
      for(PackagePart part : mediaParts){
        pictures.add(new XlsxPictureData(part));
      }
    }
    return Collections.unmodifiableList(pictures);
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public boolean isDate1904() {
    return reader.isDate1904();
  }

}
