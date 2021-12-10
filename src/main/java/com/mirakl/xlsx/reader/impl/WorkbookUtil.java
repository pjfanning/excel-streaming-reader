package com.mirakl.xlsx.reader.impl;

import com.mirakl.xlsx.reader.impl.ooxml.OoxmlReader;
import com.mirakl.xlsx.reader.XmlUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;

public class WorkbookUtil {
  public static boolean use1904Dates(OoxmlReader reader) throws IOException, InvalidFormatException, ParserConfigurationException, SAXException {
    NodeList workbookPr = XmlUtils.searchForNodeList(XmlUtils.readDocument(reader.getWorkbookData()), "/ss:workbook/ss:workbookPr");
    if (workbookPr.getLength() == 1) {
      final Node date1904 = workbookPr.item(0).getAttributes().getNamedItem("date1904");
      if (date1904 != null) {
        String value = date1904.getTextContent();
        return XmlUtils.evaluateBoolean(value);
      }
    }
    return false;
  }
}
