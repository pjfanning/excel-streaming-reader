package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.poi.xssf.streaming.SharedStringsTableBase;
import com.github.pjfanning.xlsx.SharedFormula;
import com.github.pjfanning.xlsx.StreamingReader;
import com.github.pjfanning.xlsx.XmlUtils;
import com.github.pjfanning.xlsx.exceptions.CloseException;
import com.github.pjfanning.xlsx.exceptions.NotSupportedException;
import com.github.pjfanning.xlsx.exceptions.ParseException;
import com.github.pjfanning.xlsx.impl.ooxml.HyperlinkData;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPane;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.*;

public class StreamingSheetReader implements Iterable<Row> {
  private static final Logger LOG = LoggerFactory.getLogger(StreamingSheetReader.class);
  private static final QName QNAME_HIDDEN = QName.valueOf("hidden");
  private static final QName QNAME_HT = QName.valueOf("ht");
  private static final QName QNAME_MAX = QName.valueOf("max");
  private static final QName QNAME_MIN = QName.valueOf("min");
  private static final QName QNAME_R = QName.valueOf("r");
  private static final QName QNAME_REF = QName.valueOf("ref");
  private static final QName QNAME_S = QName.valueOf("s");
  private static final QName QNAME_T = QName.valueOf("t");
  private static final QName QNAME_WIDTH = QName.valueOf("width");
  private static XMLInputFactory xmlInputFactory;

  private final StreamingWorkbookReader streamingWorkbookReader;
  private final PackagePart packagePart;
  private final SharedStrings sst;
  private final StylesTable stylesTable;
  private final Comments commentsTable;
  private final XMLEventReader parser;
  private final DataFormatter dataFormatter = new DataFormatter();
  private final Set<Integer> hiddenColumns = new HashSet<>();
  private final Map<Integer, Float> columnWidths = new HashMap<>();
  private final List<CellRangeAddress> mergedCells = new ArrayList<>();
  private final List<HyperlinkData> hyperlinks = new ArrayList<>();
  private List<XlsxHyperlink> xlsxHyperlinks;
  private Map<String, SharedFormula> sharedFormulaMap;

  private int firstRowNum = 0;
  private int lastRowNum;
  private int currentRowNum;
  private int firstColNum = 0;
  private int currentColNum;
  private final int rowCacheSize;
  private float defaultRowHeight = 0.0f;
  private int baseColWidth = 8; //POI XSSFSheet default
  private List<Row> rowCache = new ArrayList<>();
  private Iterator<Row> rowCacheIterator;

  private StringBuilder contentBuilder = new StringBuilder(64);
  private StringBuilder formulaBuilder = new StringBuilder(64);
  private StreamingSheet sheet;
  private StreamingRow currentRow;
  private StreamingCell currentCell;
  private CellAddress activeCell;
  private final boolean use1904Dates;
  private boolean insideCharElement = false;
  private boolean insideFormulaElement = false;
  private boolean insideIS = false;
  private PaneInformation pane;

  StreamingSheetReader(final StreamingWorkbookReader streamingWorkbookReader,
                       final PackagePart packagePart,
                       final SharedStrings sst, final StylesTable stylesTable, final Comments commentsTable,
                       final boolean use1904Dates, final int rowCacheSize) throws IOException, XMLStreamException {
    this.streamingWorkbookReader = streamingWorkbookReader;
    this.packagePart = packagePart;
    this.sst = sst;
    this.stylesTable = stylesTable;
    this.commentsTable = commentsTable;
    this.parser = getXmlInputFactory().createXMLEventReader(packagePart.getInputStream());

    this.use1904Dates = use1904Dates;
    this.rowCacheSize = rowCacheSize;
  }

  void setSheet(StreamingSheet sheet) {
    this.sheet = sheet;
  }

  Map<String, SharedFormula> getSharedFormulaMap() {
    if (getBuilder().readSharedFormulas()) {
      if (sharedFormulaMap == null) {
        return Collections.emptyMap();
      }
      return Collections.unmodifiableMap(sharedFormulaMap);
    } else {
      throw new IllegalStateException("The reading of shared formulas has been disabled. Enable using StreamingReader.Builder.");
    }
  }

  void addSharedFormula(String siValue, SharedFormula sharedFormula) {
    if (getBuilder().readSharedFormulas()) {
      if (sharedFormulaMap == null) {
        sharedFormulaMap = new HashMap<>();
      }
      sharedFormulaMap.put(siValue, sharedFormula);
    }
  }

  SharedFormula removeSharedFormula(String siValue) {
    if (sharedFormulaMap != null) {
      return sharedFormulaMap.remove(siValue);
    }
    return null;
  }

  boolean isUse1904Dates() {
    return use1904Dates;
  }

  float getDefaultRowHeight() {
    return defaultRowHeight;
  }

  int getBaseColWidth() {
    return baseColWidth;
  }

  /**
   * Read through a number of rows equal to the rowCacheSize field or until there is no more data to read
   *
   * @return true if data was read
   */
  private boolean getRow() {
    try {
      rowCache.clear();
      while(rowCache.size() < rowCacheSize && parser.hasNext()) {
        handleEvent(parser.nextEvent());
      }
      rowCacheIterator = rowCache.iterator();
      return rowCacheIterator.hasNext();
    } catch(XMLStreamException e) {
      throw new ParseException("Error reading XML stream", e);
    }
  }

  private void handleEvent(XMLEvent event) {
    if (event.getEventType() == XMLStreamConstants.CHARACTERS) {
      if (insideCharElement) {
        contentBuilder.append(event.asCharacters().getData());
      }
      if (insideFormulaElement) {
        formulaBuilder.append(event.asCharacters().getData());
      }
    } else if (event.getEventType() == XMLStreamConstants.START_ELEMENT
            && isSpreadsheetTag(event.asStartElement().getName())) {
      StartElement startElement = event.asStartElement();
      String tagLocalName = startElement.getName().getLocalPart();

      if ("row".equals(tagLocalName)) {
        Attribute rowNumAttr = startElement.getAttributeByName(QNAME_R);
        int rowIndex = currentRowNum;
        if (rowNumAttr != null) {
          rowIndex = Integer.parseInt(rowNumAttr.getValue()) - 1;
          currentRowNum = rowIndex;
        }
        Attribute isHiddenAttr = startElement.getAttributeByName(QNAME_HIDDEN);
        Attribute htAttr = startElement.getAttributeByName(QNAME_HT);
        float height = getDefaultRowHeight();
        if (htAttr != null) {
          try {
            height = Float.parseFloat(htAttr.getValue());
          } catch (Exception e) {
            LOG.warn("unable to parse row {} height {}", rowIndex, htAttr.getValue());
          }
        }
        boolean isHidden = isHiddenAttr != null && XmlUtils.evaluateBoolean(isHiddenAttr.getValue());
        currentRow = new StreamingRow(sheet, rowIndex, isHidden);
        currentRow.setStreamingSheetReader(this);
        currentRow.setHeight(height);
        currentColNum = firstColNum;
      } else if ("col".equals(tagLocalName)) {
        Attribute isHiddenAttr = startElement.getAttributeByName(QNAME_HIDDEN);
        Attribute widthAttr = startElement.getAttributeByName(QNAME_WIDTH);
        float width = -1;
        if (widthAttr != null) {
          try {
            width = Float.parseFloat(widthAttr.getValue());
          } catch (Exception e) {
            LOG.warn("Failed to parse column width {}", width);
          }
        }
        boolean isHidden = isHiddenAttr != null && XmlUtils.evaluateBoolean(isHiddenAttr.getValue());
        if (isHidden || width >= 0) {
          Attribute minAttr = startElement.getAttributeByName(QNAME_MIN);
          Attribute maxAttr = startElement.getAttributeByName(QNAME_MAX);
          int min = Integer.parseInt(minAttr.getValue()) - 1;
          int max = Integer.parseInt(maxAttr.getValue()) - 1;
          for (int columnIndex = min; columnIndex <= max; columnIndex++) {
            if (isHidden) hiddenColumns.add(columnIndex);
            if (width >= 0) columnWidths.put(columnIndex, width);
          }
        }
      } else if ("c".equals(tagLocalName)) {
        Attribute ref = startElement.getAttributeByName(QNAME_R);

        if (ref != null) {
          CellAddress cellAddress = new CellAddress(ref.getValue());
          currentColNum = cellAddress.getColumn();
          if (currentRow.getRowNum() == currentRowNum) {
            currentCell = new StreamingCell(sheet, currentColNum, currentRow, use1904Dates);
          } else {
            currentCell = new StreamingCell(sheet, currentColNum, cellAddress.getRow(), use1904Dates);
          }
        } else if (currentRow != null) {
          currentCell = new StreamingCell(sheet, currentColNum, currentRow, use1904Dates);
        } else {
          currentCell = new StreamingCell(sheet, currentColNum, currentRowNum, use1904Dates);
        }
        setFormatString(startElement, currentCell);

        Attribute type = startElement.getAttributeByName(QNAME_T);
        if (type != null) {
          currentCell.setType(type.getValue());
        } else {
          currentCell.setType("n");
        }

        if (stylesTable != null) {
          Attribute style = startElement.getAttributeByName(QNAME_S);
          if (style != null) {
            String indexStr = style.getValue();
            try {
              int index = Integer.parseInt(indexStr);
              currentCell.setCellStyle(stylesTable.getStyleAt(index));
            } catch (NumberFormatException nfe) {
              LOG.warn("Ignoring invalid style index {}", indexStr);
            }
          } else {
            currentCell.setCellStyle(stylesTable.getStyleAt(0));
          }
        }
      } else if ("pane".equals(tagLocalName)) {
        parsePane(startElement);
      } else if ("v".equals(tagLocalName) || "t".equals(tagLocalName)) {
        insideCharElement = true;
      } else if ("is".equals(tagLocalName)) {
        insideIS = true;
      } else if ("dimension".equals(tagLocalName)) {
        Attribute refAttr = startElement.getAttributeByName(QNAME_REF);
        String ref = refAttr != null ? refAttr.getValue() : null;
        if (ref != null) {
          // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
          for (int i = ref.length() - 1; i >= 0; i--) {
            if (!Character.isDigit(ref.charAt(i))) {
              try {
                lastRowNum = Integer.parseInt(ref.substring(i + 1)) - 1;
              } catch (NumberFormatException ignore) {
              }
              break;
            }
          }
          int colonPos = ref.indexOf(':');
          if (colonPos > 0) {
            String firstPart = ref.substring(0, colonPos);
            try {
              CellReference cellReference = new CellReference(firstPart);
              firstRowNum = cellReference.getRow();
            } catch (Exception e) {
              LOG.warn("Failed to parse cell reference {}", firstPart);
            }
          }
          for (int i = 0; i < ref.length(); i++) {
            if (!Character.isAlphabetic(ref.charAt(i))) {
              firstColNum = CellReference.convertColStringToIndex(ref.substring(0, i));
              break;
            }
          }
        }
      } else if ("f".equals(tagLocalName)) {
        insideFormulaElement = true;
        if (currentCell != null) {
          currentCell.setFormulaType(true);
          Attribute tAttr = startElement.getAttributeByName(new QName("t"));
          if (tAttr != null && tAttr.getValue().equals("shared")) {
            currentCell.setSharedFormula(true);
          }
          Attribute siAttr = startElement.getAttributeByName(new QName("si"));
          if (siAttr != null) {
            currentCell.setFormulaSI(siAttr.getValue());
          }
        }
      } else if ("mergeCell".equals(tagLocalName)) {
        parseMergeCell(startElement);
      } else if ("selection".equals(tagLocalName)) {
        Attribute activeCellAttr = startElement.getAttributeByName(QName.valueOf("activeCell"));
        if (activeCellAttr != null) {
          String activeCellRef = getAttributeValue(activeCellAttr);
          try {
            this.activeCell = new CellAddress(activeCellRef);
          } catch (Exception e) {
            LOG.warn("unable to parse active cell reference {}", activeCellRef);
          }
        }
      } else if ("hyperlink".equals(tagLocalName)) {
        parseHyperlink(startElement);
      } else if ("sheetFormatPr".equals(tagLocalName)) {
        parseSheetFormatPr(startElement);
      }

      if (!insideIS) {
        contentBuilder.setLength(0);
      }
      formulaBuilder.setLength(0);
    } else if (event.getEventType() == XMLStreamConstants.END_ELEMENT
            && isSpreadsheetTag(event.asEndElement().getName())) {
      EndElement endElement = event.asEndElement();
      String tagLocalName = endElement.getName().getLocalPart();

      if ("v".equals(tagLocalName) || "t".equals(tagLocalName)) {
        insideCharElement = false;
        Supplier formattedContentSupplier = formattedContents();
        currentCell.setRawContents(unformattedContents(formattedContentSupplier));
        currentCell.setContentSupplier(formattedContentSupplier);
      } else if ("row".equals(tagLocalName) && currentRow != null) {
        rowCache.add(currentRow);
        currentRowNum++;
      } else if ("c".equals(tagLocalName)) {
        if (currentRow == null) {
          final CellAddress cellAddress = currentCell == null ? null : currentCell.getAddress();
          LOG.warn("failed to add cell {} to cell map because currentRow is null", cellAddress);
        } else {
          currentRow.getCellMap().put(currentCell.getColumnIndex(), currentCell);
        }
        currentCell = null;
        currentColNum++;
      } else if ("is".equals(tagLocalName)) {
        insideIS = false;
      } else if ("f".equals(tagLocalName)) {
        insideFormulaElement = false;
        if (currentCell != null) {
          final String formula = formulaBuilder.toString();
          currentCell.setFormula(formula);
          if (currentCell.isSharedFormula() && currentCell.getFormulaSI() != null && getBuilder().readSharedFormulas()) {
            if (sharedFormulaMap == null) {
              sharedFormulaMap = new HashMap<>();
            }
            if (!sharedFormulaMap.containsKey(currentCell.getFormulaSI()) && !formula.isEmpty()) {
              sharedFormulaMap.put(currentCell.getFormulaSI(), new SharedFormula(currentCell.getAddress(), formula));
            } else if (formula.isEmpty()) {
              Workbook wb = getWorkbook();
              if (wb != null) {
                SharedFormula sf = sharedFormulaMap.get(currentCell.getFormulaSI());
                if (sf == null) {
                  LOG.warn("No SharedFormula found for si={}", currentCell.getFormulaSI());
                } else {
                  CurrentRowEvaluationWorkbook evaluationWorkbook =
                          new CurrentRowEvaluationWorkbook(wb, currentRow);
                  int sheetIndex = wb.getSheetIndex(sheet);
                  if (sheetIndex < 0) {
                    LOG.warn("Failed to find correct sheet index; defaulting to zero");
                    sheetIndex = 0;
                  }
                  try {
                    Ptg[] ptgs = FormulaParser.parse(sf.getFormula(), evaluationWorkbook, FormulaType.CELL, sheetIndex, currentRow.getRowNum());
                    String shiftedFmla = null;
                    final int rowsToMove = currentRowNum - sf.getCellAddress().getRow();
                    FormulaShifter formulaShifter = FormulaShifter.createForRowShift(sheetIndex, sheet.getSheetName(),
                            0, SpreadsheetVersion.EXCEL2007.getLastRowIndex(), rowsToMove, SpreadsheetVersion.EXCEL2007);
                    if (formulaShifter.adjustFormula(ptgs, sheetIndex)) {
                      shiftedFmla = FormulaRenderer.toFormulaString(evaluationWorkbook, ptgs);
                    }
                    LOG.debug("cell {} should have formula {} based on shared formula {} (rowsToMove={})",
                            currentCell.getAddress(), shiftedFmla, sf.getFormula(), rowsToMove);
                    currentCell.setFormula(shiftedFmla);
                  } catch (Exception e) {
                    LOG.warn("cell {} should has a shared formula but excel-streaming-reader has an issue parsing it - will ignore the formula",
                            currentCell.getAddress(), e);
                  }
                }
              }
            } else {
              LOG.error("No eval workbook found");
            }
          }
        }
      }
    }
  }

  private void parseHyperlink(StartElement startElement) {
    String id = null;
    Iterator<Attribute> attributeIterator = startElement.getAttributes();
    while (attributeIterator.hasNext()) {
      Attribute att = attributeIterator.next();
      QName qn = att.getName();
      if ("id".equals(qn.getLocalPart()) && qn.getNamespaceURI().endsWith("relationships")) {
        id = att.getValue();
      }
    }
    Attribute ref = startElement.getAttributeByName(QNAME_REF);
    Attribute location = startElement.getAttributeByName(QName.valueOf("location"));
    Attribute display = startElement.getAttributeByName(QName.valueOf("display"));
    Attribute tooltip = startElement.getAttributeByName(QName.valueOf("tooltip"));
    hyperlinks.add(new HyperlinkData(id, getAttributeValue(ref), getAttributeValue(location),
            getAttributeValue(display), getAttributeValue(tooltip)));
  }

  private void parseMergeCell(StartElement startElement) {
    Attribute ref = startElement.getAttributeByName(QNAME_REF);
    if (ref != null) {
      mergedCells.add(CellRangeAddress.valueOf(ref.getValue()));
    }
  }

  private void parsePane(final StartElement startElement) {
    final Attribute stateAtt = startElement.getAttributeByName(QName.valueOf("state"));
    final Attribute activePaneAtt = startElement.getAttributeByName(QName.valueOf("activePane"));
    final Attribute topLeftCellAtt = startElement.getAttributeByName(QName.valueOf("topLeftCell"));
    final Float xValue = parseAttValueAsFloat("xSplit", startElement);
    final short x = xValue == null ? 0 : xValue.shortValue();
    final Float yValue = parseAttValueAsFloat("ySplit", startElement);
    final short y = yValue == null ? 0 : yValue.shortValue();
    short row = 0;
    short col = 0;
    if (topLeftCellAtt != null) {
      try {
        final CellReference cellRef = new CellReference(topLeftCellAtt.getValue());
        row = (short)cellRef.getRow();
        col = cellRef.getCol();
      } catch (Exception e) {
        LOG.warn("unable to parse topLeftCell {}", topLeftCellAtt.getValue());
      }
    }
    final boolean frozen = stateAtt != null && "frozen".equals(stateAtt.getValue());
    byte active = 0;
    if (activePaneAtt != null) {
      try {
        STPane.Enum stPaneEnum = STPane.Enum.forString(activePaneAtt.getValue());
        active = (byte)(stPaneEnum.intValue() - 1);
      } catch (Exception e) {
        LOG.warn("unable to parse activePane {}", activePaneAtt.getValue());
      }
    }
    pane = new PaneInformation(x, y, row, col, active, frozen);
  }

  private Float parseAttValueAsFloat(final String name, final StartElement startElement) {
    final Attribute att = startElement.getAttributeByName(QName.valueOf(name));
    if (att != null) {
      try {
        return Float.parseFloat(att.getValue());
      } catch (Exception e) {
        LOG.warn("unable to parse {} {}", name, att.getValue());
      }
    }
    return null;
  }

  private void parseSheetFormatPr(final StartElement startElement) {
    final Attribute defaultRowHeightAtt = startElement.getAttributeByName(QName.valueOf("defaultRowHeight"));
    if (defaultRowHeightAtt != null) {
      try {
        defaultRowHeight = Float.parseFloat(defaultRowHeightAtt.getValue());
      } catch (Exception e) {
        LOG.warn("unable to parse defaultRowHeight {}", defaultRowHeightAtt.getValue());
      }
    }
    final Attribute baseColWidthAtt = startElement.getAttributeByName(QName.valueOf("baseColWidth"));
    if (baseColWidthAtt != null) {
      try {
        baseColWidth = Integer.parseInt(baseColWidthAtt.getValue());
      } catch (Exception e) {
        LOG.warn("unable to parse baseColWidth {}", baseColWidthAtt.getValue());
      }
    }
  }

  /**
   * Returns true if a tag is part of the main namespace for SpreadsheetML:
   * <ul>
   * <li>http://schemas.openxmlformats.org/spreadsheetml/2006/main
   * <li>http://purl.oclc.org/ooxml/spreadsheetml/main
   * </ul>
   * As opposed to http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing, etc.
   */
  private boolean isSpreadsheetTag(QName name) {
    return (name.getNamespaceURI() != null
        && name.getNamespaceURI().endsWith("/main"));
  }

  /**
   * Get the hidden state for a given column
   *
   * @param columnIndex - the column to set (0-based)
   * @return hidden - <code>false</code> if the column is visible
   */
  boolean isColumnHidden(int columnIndex) {
    if(rowCacheIterator == null) {
      getRow();
    }
    return hiddenColumns.contains(columnIndex);
  }

  float getColumnWidth(int columnIndex) {
    if(rowCacheIterator == null) {
      getRow();
    }
    Float width = columnWidths.get(columnIndex);
    return width == null ? getBaseColWidth() : width;
  }

  /**
   * Gets the last row on the sheet
   */
  int getFirstRowNum() {
    if(rowCacheIterator == null) {
      getRow();
    }
    return firstRowNum;
  }

  /**
   * Gets the last row on the sheet
   */
  int getLastRowNum() {
    if(rowCacheIterator == null) {
      getRow();
    }
    return lastRowNum;
  }

  /**
   * Read the numeric format string out of the styles table for this cell. Stores
   * the result in the Cell.
   *
   * @param startElement
   * @param cell
   */
  void setFormatString(StartElement startElement, StreamingCell cell) {
    Attribute cellStyle = startElement.getAttributeByName(new QName("s"));
    String cellStyleString = (cellStyle != null) ? cellStyle.getValue() : null;
    XSSFCellStyle style = null;

    if (stylesTable != null) {
      if(cellStyleString != null) {
        style = stylesTable.getStyleAt(Integer.parseInt(cellStyleString));
      } else if(stylesTable.getNumCellStyles() > 0) {
        style = stylesTable.getStyleAt(0);
      }
    }

    if(style != null) {
      cell.setNumericFormatIndex(style.getDataFormat());
      String formatString = style.getDataFormatString();

      if(formatString != null) {
        cell.setNumericFormat(formatString);
      } else {
        cell.setNumericFormat(BuiltinFormats.getBuiltinFormat(cell.getNumericFormatIndex()));
      }
    } else {
      cell.setNumericFormatIndex(null);
      cell.setNumericFormat(null);
    }
  }

  CellAddress getActiveCell() {
    return activeCell;
  }

  PaneInformation getPane() {
    if(rowCacheIterator == null) {
      getRow();
    }
    return pane;
  }

  /**
   * Tries to format the contents of the last contents appropriately based on
   * the type of cell and the discovered numeric format.
   */
  private Supplier formattedContents() {
    return getFormatterForType(currentCell.getType());
  }

  /**
   * Tries to format the contents of the last contents appropriately based on
   * the provided type and the discovered numeric format.
   */
  private Supplier getFormatterForType(String type) {
    final String lastContents = contentBuilder.toString();
    switch(type) {
      case "s":           //string stored in shared table
        if (!lastContents.isEmpty()) {
          int idx = Integer.parseInt(lastContents);
          if (!getBuilder().fullFormatRichText() && sst instanceof SharedStringsTableBase) {
            return new StringSupplier(((SharedStringsTableBase)sst).getString(idx));
          }
          return new RichTextStringSupplier(sst.getItemAt(idx));
        }
        return new StringSupplier(lastContents);
      case "inlineStr":   //inline string (not in sst)
      case "str":
        return new StringSupplier(lastContents);
      case "e":           //error type
        return new StringSupplier("ERROR:  " + lastContents);
      case "n":           //numeric type
        if(currentCell.getNumericFormat() != null && lastContents.length() > 0) {
          // the formatRawCellContents operation incurs a significant overhead on large sheets,
          // and we want to defer the execution of this method until the value is actually needed.
          // it is not needed in all cases..
          final String currentLastContents = lastContents;
          final int currentNumericFormatIndex = currentCell.getNumericFormatIndex();
          final String currentNumericFormat = currentCell.getNumericFormat();

          return new Supplier() {
            String cachedContent;

            @Override
            public Object getContent() {
              if (cachedContent == null) {
                cachedContent = dataFormatter.formatRawCellContents(
                        Double.parseDouble(currentLastContents),
                        currentNumericFormatIndex,
                        currentNumericFormat);
              }

              return cachedContent;
            }
          };
        } else {
          return new StringSupplier(lastContents);
        }
      case "d":           //date type (Strict OOXML format)
        if(currentCell.getNumericFormat() != null && lastContents.length() > 0) {
          // the formatRawCellContents operation incurs a significant overhead on large sheets,
          // and we want to defer the execution of this method until the value is actually needed.
          // it is not needed in all cases..
          final String currentLastContents = lastContents;
          final int currentNumericFormatIndex = currentCell.getNumericFormatIndex();
          final String currentNumericFormat = currentCell.getNumericFormat();

          return new Supplier() {
            String cachedContent;

            @Override
            public Object getContent() {
              if (cachedContent == null) {
                try {
                  Double dv;
                  try {
                    LocalDateTime dt = DateTimeUtil.parseDateTime(currentLastContents);
                    dv = DateUtil.getExcelDate(dt, use1904Dates);
                  } catch (Exception e) {
                    dv = DateTimeUtil.convertTime(currentLastContents);
                  }
                  cachedContent = dataFormatter.formatRawCellContents(
                          dv,
                          currentNumericFormatIndex,
                          currentNumericFormat);
                } catch (Exception e) {
                  LOG.warn("cannot format strict format date/time {}", currentLastContents);
                  cachedContent = currentLastContents;
                }
              }
              return cachedContent;
            }
          };
        } else {
          return new StringSupplier(lastContents);
        }
      default:
        return new StringSupplier(lastContents);
    }
  }

  /**
   * Returns the contents of the cell, with no formatting applied
   */
  private String unformattedContents(Supplier formattedContentSupplier) {
    final String lastContents = contentBuilder.toString();
    switch(currentCell.getType()) {
      case "s":           //string stored in shared table
        Object formattedContent = formattedContentSupplier.getContent();
        if (formattedContent instanceof RichTextString) {
          return ((RichTextString) formattedContent).getString();
        } else if (formattedContent != null) {
          return formattedContent.toString();
        }
        if (!lastContents.isEmpty()) {
          int idx = Integer.parseInt(lastContents);
          if (sst == null) throw new NullPointerException("sst is null");
          if (sst instanceof SharedStringsTableBase) {
            return ((SharedStringsTableBase)sst).getString(idx);
          }
          return sst.getItemAt(idx).getString();
        }
        return lastContents;
      case "inlineStr":   //inline string (not in sst)
        return new XSSFRichTextString(lastContents).getString();
      default:
        return lastContents;
    }
  }

  /**
   * Returns a new streaming iterator to loop through rows. This iterator is not
   * guaranteed to have all rows in memory, and any particular iteration may
   * trigger a load from disk to read in new data.
   *
   * @return the streaming iterator
   */
  @Override
  public Iterator<Row> iterator() {
    return new StreamingRowIterator();
  }

  /**
   * @return the comments associated with this sheet (only if feature is enabled on the Builder)
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadComments(boolean)} is not set to true
   */
  Comments getCellComments() {
    if (!streamingWorkbookReader.getBuilder().readComments()) {
      throw new IllegalStateException("getCellComments() only works if StreamingWorking.Builder setReadComments is set to true");
    }
    return this.commentsTable;
  }

  List<CellRangeAddress> getMergedCells() { return this.mergedCells; }

  XSSFDrawing getDrawingPatriarch() {
    if (!streamingWorkbookReader.getBuilder().readShapes()) {
      throw new IllegalStateException("getDrawingPatriarch() only works if StreamingWorking.Builder setReadShapes is set to true");
    }
    if (sheet != null) {
      List<XSSFShape> shapes = streamingWorkbookReader.getShapes(sheet.getSheetName());
      if (shapes != null) {
        Iterator<XSSFShape> shapesIter = shapes.iterator();
        while (shapesIter.hasNext()) {
          return shapesIter.next().getDrawing();
        }
      }
    }
    return null;
  }

  public void close() {
    try {
      parser.close();
    } catch(XMLStreamException e) {
      throw new CloseException(e);
    }
  }

  StreamingReader.Builder getBuilder() {
    return streamingWorkbookReader.getBuilder();
  }

  Workbook getWorkbook() {
    return streamingWorkbookReader.getWorkbook();
  }

  /**
   * @return the hyperlinks associated with this sheet (only if feature is enabled on the Builder)
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  List<XlsxHyperlink> getHyperlinks() {
    if (!getBuilder().readHyperlinks()) {
      throw new IllegalStateException("getHyperlinks() only works if StreamingWorking.Builder setReadHyperlinks is set to true");
    }
    initHyperlinks();
    return xlsxHyperlinks;
  }

  private String getAttributeValue(Attribute att) {
    return att == null ? null : att.getValue();
  }

  private void initHyperlinks() {
    if (xlsxHyperlinks == null || xlsxHyperlinks.isEmpty()) {
      ArrayList<XlsxHyperlink> links = new ArrayList<>();

      try {
        PackageRelationshipCollection hyperRels =
                packagePart.getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation());

        // Turn each one into a XSSFHyperlink
        for(HyperlinkData hyperlink : hyperlinks) {
          PackageRelationship hyperRel = null;
          if(hyperlink.getId() != null) {
            hyperRel = hyperRels.getRelationshipByID(hyperlink.getId());
          }

          links.add( new XlsxHyperlink(hyperlink, hyperRel) );
        }
      } catch (InvalidFormatException e){
        throw new POIXMLException(e);
      }
      xlsxHyperlinks = links;
    }
  }

  class StreamingRowIterator implements Iterator<Row> {
    public StreamingRowIterator() {
      if(rowCacheIterator == null) {
        if(!hasNext()) {
          LOG.debug("there appear to be no rows");
        }
      }
    }

    @Override
    public boolean hasNext() {
      return (rowCacheIterator != null && rowCacheIterator.hasNext()) || getRow();
    }

    @Override
    public Row next() {
      try {
        return rowCacheIterator.next();
      } catch(NoSuchElementException nsee) {
        //see https://github.com/monitorjbl/excel-streaming-reader/issues/176
        if (hasNext()) {
          return rowCacheIterator.next();
        }
        throw nsee;
      }
    }

    @Override
    public void remove() {
      throw new NotSupportedException();
    }
  }

  private static XMLInputFactory getXmlInputFactory() {
    if (xmlInputFactory == null) {
      try {
        xmlInputFactory = XMLHelper.newXMLInputFactory();
      } catch (Exception e) {
        LOG.error("Issue creating XMLInputFactory", e);
        throw e;
      }
    }
    return xmlInputFactory;
  }
}
