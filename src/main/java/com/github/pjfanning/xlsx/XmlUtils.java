package com.github.pjfanning.xlsx;

import com.github.pjfanning.xlsx.exceptions.ParseException;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.XMLConstants;
import javax.xml.namespace.NamespaceContext;
import javax.xml.xpath.*;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public final class XmlUtils {

  private static final NamespaceContext transitionalFormatNamespaceContext =
          new NamespaceContextImpl("ss", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
  private static final NamespaceContextImpl strictFormatNamespaceContext =
          new NamespaceContextImpl("ss", "http://purl.oclc.org/ooxml/spreadsheetml/main");

  public static final String FALSE_AS_STRING = "0";
  public static final String TRUE_AS_STRING  = "1";

  private XmlUtils() {}

  public static Document readDocument(InputStream inp) throws IOException, SAXException {
    return DocumentHelper.readDocument(inp);
  }

  public static NodeList searchForNodeList(Document document, String xpath) {
    try {
      XPath xp = XPathFactory.newInstance().newXPath();
      xp.setNamespaceContext(transitionalFormatNamespaceContext);
      NodeList nl = (NodeList)xp.compile(xpath).evaluate(document, XPathConstants.NODESET);
      if (nl.getLength() == 0) {
        xp.setNamespaceContext(strictFormatNamespaceContext);
        nl = (NodeList)xp.compile(xpath).evaluate(document, XPathConstants.NODESET);
      }
      return nl;
    } catch(XPathExpressionException e) {
      throw new ParseException(e);
    }
  }

  public static boolean evaluateBoolean(String bool) {
    return bool.equals(TRUE_AS_STRING) || bool.equalsIgnoreCase("true");
  }

  private static final class NamespaceContextImpl implements NamespaceContext {
    private final Map<String, String> urisByPrefix = new HashMap<>();

    private final Map<String, Set<String>> prefixesByURI = new HashMap<>();

    public NamespaceContextImpl() {
      addNamespace(XMLConstants.XML_NS_PREFIX, XMLConstants.XML_NS_URI);
      addNamespace(XMLConstants.XMLNS_ATTRIBUTE, XMLConstants.XMLNS_ATTRIBUTE_NS_URI);
    }

    public NamespaceContextImpl(String prefix, String uri) {
      this();
      addNamespace(prefix, uri);
    }

    private void addNamespace(String prefix, String namespaceURI) {
      urisByPrefix.put(prefix, namespaceURI);
      if (prefixesByURI.containsKey(namespaceURI)) {
        (prefixesByURI.get(namespaceURI)).add(prefix);
      } else {
        Set<String> set = new HashSet<>();
        set.add(prefix);
        prefixesByURI.put(namespaceURI, set);
      }
    }

    @Override
    public String getNamespaceURI(String prefix) {
      if (prefix == null)
        throw new IllegalArgumentException("prefix cannot be null");
      return urisByPrefix.getOrDefault(prefix, XMLConstants.NULL_NS_URI);
    }

    @Override
    public String getPrefix(String namespaceURI) {
      return getPrefixes(namespaceURI).next();
    }

    @Override
    public Iterator<String> getPrefixes(String namespaceURI) {
      if (namespaceURI == null)
        throw new IllegalArgumentException("namespaceURI cannot be null");
      if (prefixesByURI.containsKey(namespaceURI)) {
        return (prefixesByURI.get(namespaceURI)).iterator();
      } else {
        return Collections.emptyIterator();
      }
    }
  }
}
