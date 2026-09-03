package com.github.pjfanning.xlsx.impl;

import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.util.ArrayList;
import java.util.List;

/**
 * Small DOM traversal helpers used instead of XPath for the couple of trivial lookups this
 * library needs in workbook.xml.
 * <p>
 * XPathFactory is documented as not thread-safe, so it cannot be cached in a static field, and
 * creating one per call (plus compiling the expression) costs far more than walking a handful of
 * child nodes. Matching on local name also removes the need to retry the lookup against the
 * strict OOXML namespace.
 * </p>
 */
final class DomUtil {

  private DomUtil() {}

  /**
   * @param parent the node whose direct children should be searched (can be null)
   * @param localName the local name to match, ignoring namespace
   * @return the matching direct child elements, never null
   */
  static List<Element> getChildElements(final Node parent, final String localName) {
    final List<Element> elements = new ArrayList<>();
    if (parent == null) {
      return elements;
    }
    final NodeList children = parent.getChildNodes();
    for (int i = 0; i < children.getLength(); i++) {
      final Node child = children.item(i);
      if (child instanceof Element && localName.equals(getLocalName(child))) {
        elements.add((Element) child);
      }
    }
    return elements;
  }

  /**
   * @param parent the node whose direct children should be searched (can be null)
   * @param localName the local name to match, ignoring namespace
   * @return the first matching direct child element, or null if there is none
   */
  static Element getFirstChildElement(final Node parent, final String localName) {
    final List<Element> elements = getChildElements(parent, localName);
    return elements.isEmpty() ? null : elements.get(0);
  }

  private static String getLocalName(final Node node) {
    //getLocalName is null for nodes created by a namespace unaware parser
    final String localName = node.getLocalName();
    return localName == null ? node.getNodeName() : localName;
  }
}
