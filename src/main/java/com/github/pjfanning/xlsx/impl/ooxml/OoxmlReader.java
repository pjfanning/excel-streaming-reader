/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package com.github.pjfanning.xlsx.impl.ooxml;

import com.github.pjfanning.xlsx.CommentsImplementationType;
import com.github.pjfanning.xlsx.SharedStringsImplementationType;
import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.exceptions.MissingSheetException;
import com.github.pjfanning.xlsx.exceptions.OpenException;
import com.github.pjfanning.xlsx.impl.PoiSharedStringsSupport;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.util.Internal;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.xmlbeans.XmlException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import javax.xml.stream.XMLStreamException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@Internal
public class OoxmlReader extends XSSFReader {

  private static final Logger LOGGER = LoggerFactory.getLogger(OoxmlReader.class);
  private final OoxmlSheetReader ooxmlSheetReader;

  /**
   * Creates a new XSSFReader, for the given package
   * @param builder the builder to use
   * @param pkg The package this poi filesystem reader will use
   * @param strictOoxmlChecksNeeded whether to perform strict OOXML checks
   * @throws IOException if an error occurs while reading the package
   * @throws OpenXML4JException if an error occurs while parsing the package
   * @throws POIXMLException if the package is invalid
   */
  @Internal
  public OoxmlReader(StreamingReader.Builder builder,
                     OPCPackage pkg, boolean strictOoxmlChecksNeeded)
          throws IOException, OpenXML4JException, POIXMLException {
    super(pkg, true);
    ooxmlSheetReader = new OoxmlSheetReader(builder, this, strictOoxmlChecksNeeded);
  }

  /**
   * Opens up the Shared Strings Table, parses it, and
   * returns a handy object for working with
   * shared strings.
   * <p>
   * Only {@link SharedStringsImplementationType#POI_DEFAULT} and
   * {@link SharedStringsImplementationType#POI_READ_ONLY} are handled here. The temp file
   * and map backed implementations are created by the workbook reader before this is called.
   * </p>
   * @param builder the builder to use
   * @return the shared strings table
   * @throws IOException if an error occurs while reading the shared strings
   * @throws SAXException if an error occurs while parsing the shared strings
   */
  public SharedStrings getSharedStrings(StreamingReader.Builder builder) throws IOException, SAXException {
    final List<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
    if (parts.isEmpty()) {
      return null;
    }
    final PackagePart part = parts.get(0);
    return builder.getSharedStringsImplementationType() == SharedStringsImplementationType.POI_DEFAULT
            ? new SharedStringsTable(part)
            : new ReadOnlySharedStringsTable(part, builder.includePhoneticRuns());
  }

  public int getSheetIndex(final String name) throws MissingSheetException {
    return ooxmlSheetReader.getSheetIndex(name);
  }

  public SheetData getSheetDataAt(final int idx) throws MissingSheetException {
    return ooxmlSheetReader.getSheetData(idx);
  }

  public int getNumberOfSheets() {
    return ooxmlSheetReader.size();
  }

  static class OoxmlSheetReader {

    public final StreamingReader.Builder builder;

    /**
     * Provides random access to the sheets, which POI returns in {@code logical} order,
     * i.e. the order of the CTSheet elements in workbook.xml
     */
    private final XSSFReader.SheetIterator sheetIterator;

    private final List<XSSFSheetRef> sheetRefList;

    private final boolean strictOoxmlChecksNeeded;

    private final Map<XSSFSheetRef, SheetData> sheetDataMap = new HashMap<>();

    /**
     * Construct a new OoxmlSheetReader
     *
     * @param reader the XSSFReader for the workbook
     * @throws IOException if reading the data from the package fails
     * @throws POIXMLException if the package data is invalid
     */
    OoxmlSheetReader(final StreamingReader.Builder builder,
                     final XSSFReader reader, final boolean strictOoxmlChecksNeeded)
            throws IOException, POIXMLException {
      this.builder = builder;
      this.strictOoxmlChecksNeeded = strictOoxmlChecksNeeded;
      try {
        sheetIterator = reader.getSheetIterator();
      } catch (InvalidFormatException e) {
        throw new POIXMLException(e);
      } catch (POIXMLException e) {
        if (e.getCause() instanceof SAXException saxException) {
          throw new OpenException(saxException);
        }
        throw e;
      }
      sheetRefList = sheetIterator.getSheetRefs();
    }

    int size() {
      return sheetRefList.size();
    }

    int getSheetIndex(final String name) throws MissingSheetException {
      int i = 0;
      for (XSSFSheetRef sheetRef : sheetRefList) {
        if (name.equalsIgnoreCase(sheetRef.getName())) {
          return i;
        }
        i++;
      }
      throw new MissingSheetException("Failed to find sheet " + name);
    }

    SheetData getSheetData(final int idx) throws MissingSheetException {
      if (idx < 0 || idx >= sheetRefList.size()) {
        throw new MissingSheetException("Failed to find sheet with id " + idx);
      }
      XSSFSheetRef matchedSheetRef = sheetRefList.get(idx);
      if (matchedSheetRef == null) {
        throw new MissingSheetException("Failed to find sheet with id " + idx);
      }
      return getSheetData(matchedSheetRef);
    }

    SheetData getSheetData(final XSSFSheetRef sheetRef) {
      return sheetDataMap.computeIfAbsent(sheetRef, ref ->
              createSheetData(builder, ref, sheetIterator.getSheetPart(ref), strictOoxmlChecksNeeded));
    }
  }

  public static class SheetData {
    private final PackagePart sheetPart;
    private final String sheetName;
    private final Comments comments;
    private final List<XSSFShape> shapes;

    SheetData(final PackagePart sheetPart, final String sheetName, final Comments comments, List<XSSFShape> shapes) {
      this.sheetPart = sheetPart;
      this.sheetName = sheetName;
      this.comments = comments;
      this.shapes = shapes;
    }

    public PackagePart getSheetPart() {
      return sheetPart;
    }

    public String getSheetName() {
      return sheetName;
    }

    public Comments getComments() {
      return comments;
    }

    public List<XSSFShape> getShapes() {
      return shapes;
    }
  }

  private static SheetData createSheetData(final StreamingReader.Builder builder,
                                           final XSSFSheetRef xssfSheetRef,
                                           final PackagePart sheetPart,
                                           final boolean strictOoxmlChecksNeeded) {
    if (sheetPart == null) {
      return new SheetData(null, xssfSheetRef.getName(), null, null);
    }
    final List<XSSFShape> shapes = builder.readShapes() ? getShapes(sheetPart, strictOoxmlChecksNeeded) : null;
    final Comments comments = builder.readComments() ? getSheetComments(builder, sheetPart, strictOoxmlChecksNeeded) : null;
    return new SheetData(sheetPart, xssfSheetRef.getName(), comments, shapes);
  }

  /**
   * Returns the comments associated with this sheet,
   * or null if there aren't any
   */
  private static Comments getSheetComments(final StreamingReader.Builder builder,
                                           final PackagePart sheetPkg,
                                           final boolean strictOoxmlChecksNeeded) {
    try {
      final PackagePart commentsPart = XSSFReader.getSheetCommentsPart(sheetPkg, strictOoxmlChecksNeeded);
      return commentsPart == null ? null : parseComments(builder, commentsPart, strictOoxmlChecksNeeded);
    } catch (InvalidFormatException | IOException | XMLStreamException e) {
      LOGGER.warn("issue getting sheet comments", e);
      return null;
    }
  }

  private static Comments parseComments(final StreamingReader.Builder builder,
                                        final PackagePart commentsPart,
                                        final boolean strictOoxmlChecksNeeded)
          throws IOException, XMLStreamException {
    if (builder.getCommentsImplementationType() == CommentsImplementationType.TEMP_FILE_BACKED) {
      try (InputStream is = commentsPart.getInputStream()) {
        Comments ct = PoiSharedStringsSupport.createTempFileCommentsTable(builder);
        try {
          PoiSharedStringsSupport.readComments(ct, is);
        } catch (IOException | RuntimeException e) {
          if (ct instanceof AutoCloseable) {
            try {
              ((AutoCloseable) ct).close();
            } catch (Exception e2) {
              e.addSuppressed(e2);
            }
          }
          throw e;
        }
        return ct;
      }
    } else if (builder.getCommentsImplementationType() == CommentsImplementationType.CUSTOM_MAP_BACKED) {
      try (InputStream is = commentsPart.getInputStream()) {
        Comments ct = PoiSharedStringsSupport.createMapBackedCommentsTable(builder);
        try {
          PoiSharedStringsSupport.readComments(ct, is);
        } catch (IOException | RuntimeException e) {
          if (ct instanceof AutoCloseable) {
            try {
              ((AutoCloseable) ct).close();
            } catch (Exception e2) {
              e.addSuppressed(e2);
            }
          }
          throw e;
        }
        return ct;
      }
    } else if (strictOoxmlChecksNeeded) {
      return OoxmlStrictHelper.getCommentsTable(builder, commentsPart);
    } else {
      return new CommentsTable(commentsPart);
    }
  }

  /**
   * Returns the shapes associated with this sheet,
   * an empty list or null if there is an exception
   */
  private static List<XSSFShape> getShapes(final PackagePart sheetPkg, final boolean strictOoxmlChecksNeeded) {
    final List<XSSFShape> shapes = new LinkedList<>();
    try {
      for (PackagePart drawingsPart : XSSFReader.getSheetDrawingParts(sheetPkg, strictOoxmlChecksNeeded)) {
        XSSFDrawing drawing = new XSSFDrawing(drawingsPart);
        shapes.addAll(drawing.getShapes());
      }
    } catch (XmlException | InvalidFormatException | IOException e) {
      LOGGER.warn("issue getting shapes", e);
      return null;
    }
    return shapes;
  }

}
