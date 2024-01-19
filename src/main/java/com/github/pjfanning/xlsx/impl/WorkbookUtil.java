package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.impl.ooxml.OoxmlReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;

import static com.github.pjfanning.xlsx.XmlUtils.*;

public class WorkbookUtil {

  private WorkbookUtil() {}

  /**
   * @param reader reader for the xlsx file
   * @return whether the stored xlsx has 1904 date format
   * @throws IOException if the workbook cannot be read
   * @throws InvalidFormatException if the workbook is invalid
   * @throws SAXException if the workbook cannot be parsed
   * @deprecated use {@link #use1904Dates(Document)}
   */
  @Deprecated
  public static boolean use1904Dates(OoxmlReader reader) throws IOException, InvalidFormatException, SAXException {
    return use1904Dates(readDocument(reader.getWorkbookData()));
  }

  /**
   * @param workbookDoc the workbook document in XML format
   * @return whether the stored xlsx has 1904 date format
   */
  public static boolean use1904Dates(Document workbookDoc) {
    NodeList workbookPr = searchForNodeList(workbookDoc, "/ss:workbook/ss:workbookPr");
    if (workbookPr.getLength() == 1) {
      final Node date1904 = workbookPr.item(0).getAttributes().getNamedItem("date1904");
      if (date1904 != null) {
        String value = date1904.getTextContent();
        return evaluateBoolean(value);
      }
    }
    return false;
  }
}
