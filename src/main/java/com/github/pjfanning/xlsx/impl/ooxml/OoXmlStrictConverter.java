package com.github.pjfanning.xlsx.impl.ooxml;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.util.Beta;
import org.apache.poi.util.XMLHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.namespace.QName;
import javax.xml.stream.*;
import javax.xml.stream.events.*;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

@Beta
public class OoXmlStrictConverter implements AutoCloseable {

    private static final Logger LOGGER = LoggerFactory.getLogger(OoXmlStrictConverter.class);
    private static final QName CONFORMANCE = new QName("conformance");
    private static final Properties mappings;
    //stringPropertyNames() builds a fresh HashSet on every call, and this is consulted once per
    //attribute, so the prefix mappings are materialised once here instead
    private static final Map<String, String> prefixMappings;

    static {
        mappings = OoXmlStrictConverterUtils.readMappings();
        final Map<String, String> prefixes = new LinkedHashMap<>();
        for (String key : mappings.stringPropertyNames()) {
            prefixes.put(key, mappings.getProperty(key));
        }
        prefixMappings = Collections.unmodifiableMap(prefixes);
    }

    private final XMLEventFactory xef;
    private final XMLEventWriter xew;
    private final XMLEventReader xer;
    private int depth = 0;
    private boolean inDateCell;
    private boolean inDateValue;

    public OoXmlStrictConverter(InputStream is, OutputStream os) throws XMLStreamException {
        this.xer = newXmlInputFactory().createXMLEventReader(is);
        this.xew = newXmlOutputFactory().createXMLEventWriter(os);
        this.xef = newXmlEventFactory();
    }

    public boolean convertNextElement() throws XMLStreamException {
        if (!xer.hasNext()) {
            return false;
        }

        XMLEvent xe = xer.nextEvent();
        if(xe.isStartElement()) {
            xew.add(convertDateStartElement(convertStartElement(xe.asStartElement(), depth==0)));
            depth++;
        } else if(xe.isEndElement()) {
            xew.add(updateDateFlagsOnEndElement(convertEndElement(xe.asEndElement())));
            depth--;
        } else {
            if (inDateValue) {
                xew.add(convertDateValueToNumeric(xe));
            } else {
                // Add as is
                xew.add(xe);
            }
        }

        return true;
    }

    private XMLEvent convertDateValueToNumeric(XMLEvent xe) {
        if (!xe.isCharacters()) {
            return xe;
        }

        Date date = DateUtil.parseYYYYMMDDDate(xe.asCharacters().getData());

        double excelDate = DateUtil.getExcelDate(date);

        return xef.createCharacters(Double.toString(excelDate));
    }

    private EndElement updateDateFlagsOnEndElement(EndElement endElement) {
        if (inDateValue) {
            if ("v".equals(endElement.getName().getLocalPart())) {
                inDateValue = false;
            }
            return endElement;
        }

        if (inDateCell) {
            if (isCell(endElement.getName())) {
                inDateCell = false;
            }
            return endElement;
        }

        return endElement;
    }

    private StartElement convertDateStartElement(StartElement startElement) {

        if (inDateCell) {
            if ("v".equals(startElement.getName().getLocalPart())) {
                this.inDateValue = true;
            }
            return startElement;
        }

        if (!isDateCell(startElement)) {
            return startElement;
        }

        this.inDateCell = true;

        // Change to numeric cell.
        return xef.createStartElement(startElement.getName(),
                changeTypeAttributeToNumeric(startElement.getAttributes()),
                startElement.getNamespaces());

    }

    private Iterator<? extends Attribute> changeTypeAttributeToNumeric(
            Iterator<Attribute> attributes) {
        List<Attribute> result = new ArrayList<>();

        while (attributes.hasNext()) {
            Attribute attribute = attributes.next();
            if (!"t".equals(attribute.getName().getLocalPart())) {
                result.add(attribute);
                continue;
            }

            result.add(xef.createAttribute(attribute.getName(), "n"));
        }

        return Collections.unmodifiableList(result).iterator();
    }

    private boolean isDateCell(StartElement startElement) {
        if (!isCell(startElement.getName())) {
            return false;
        }

        Attribute typeAttribute = startElement.getAttributeByName(QName.valueOf("t"));
        if (typeAttribute == null) {
            return false;
        }

        return "d".equals(typeAttribute.getValue());
    }

    private boolean isCell(QName elementName) {
        return "c".equals(elementName.getLocalPart());
    }


    @Override
    public void close() throws XMLStreamException {
        //the writer is flushed here rather than after every single element. All callers use
        //try-with-resources with the converter declared last, so this runs before the
        //underlying OutputStream is closed.
        try {
            xew.flush();
        } finally {
            try {
                xer.close();
            } finally {
                xew.close();
            }
        }
    }

    private StartElement convertStartElement(StartElement startElement, boolean root) {
        return xef.createStartElement(updateQName(startElement.getName()),
                processAttributes(startElement.getAttributes(), startElement.getName().getNamespaceURI(), root),
                processNamespaces(startElement.getNamespaces()));
    }

    private EndElement convertEndElement(EndElement endElement) {
        return xef.createEndElement(updateQName(endElement.getName()),
                processNamespaces(endElement.getNamespaces()));

    }

    private static QName updateQName(QName qn) {
        String namespaceUri = qn.getNamespaceURI();
        if(OoXmlStrictConverterUtils.isNotBlank(namespaceUri)) {
            String mappedUri = mappings.getProperty(namespaceUri);
            if(mappedUri != null) {
                qn = OoXmlStrictConverterUtils.isBlank(qn.getPrefix()) ? new QName(mappedUri, qn.getLocalPart())
                        : new QName(mappedUri, qn.getLocalPart(), qn.getPrefix());
            }
        }
        return qn;
    }

    private Iterator<Attribute> processAttributes(final Iterator<Attribute> iter,
            final String elementNamespaceUri, final boolean rootElement) {
        ArrayList<Attribute> list = new ArrayList<>();
        while(iter.hasNext()) {
            Attribute att = iter.next();
            QName qn = updateQName(att.getName());
            if(rootElement && mappings.containsKey(elementNamespaceUri) && att.getName().equals(CONFORMANCE)) {
                //drop attribute
            } else {
                String newValue = att.getValue();
                for(Map.Entry<String, String> mapping : prefixMappings.entrySet()) {
                    if(att.getValue().startsWith(mapping.getKey())) {
                        newValue = att.getValue().replace(mapping.getKey(), mapping.getValue());
                        break;
                    }
                }
                list.add(xef.createAttribute(qn, newValue));
            }
        }
        return Collections.unmodifiableList(list).iterator();
    }

    private Iterator<Namespace> processNamespaces(final Iterator<Namespace> iter) {
        ArrayList<Namespace> list = new ArrayList<>();
        while(iter.hasNext()) {
            Namespace ns = iter.next();
            final String mappedUri = prefixMappings.get(ns.getNamespaceURI());
            if(mappedUri != null) {
                //rewrite the declaration to the mapped namespace rather than dropping it. If it is
                //dropped, the writer is left with elements in a namespace nothing declares, and
                //POI's XMLOutputFactory has isRepairingNamespaces set, so it has to synthesise a
                //declaration for every single element.
                list.add(ns.isDefaultNamespaceDeclaration()
                        ? xef.createNamespace(mappedUri)
                        : xef.createNamespace(ns.getPrefix(), mappedUri));
            } else if(!ns.isDefaultNamespaceDeclaration()) {
                list.add(ns);
            }
        }
        return Collections.unmodifiableList(list).iterator();
    }

    /**
     * The StAX factories are created per converter instead of being cached in static fields.
     * None of them are thread-safe (the JDK XMLInputFactory writes to its own fTempReader and
     * fPropertyChanged fields on every createXMLEventReader call, and XMLEventFactory carries a
     * mutable Location), and converters can be run concurrently.
     */
    private static XMLInputFactory newXmlInputFactory() {
        try {
            return XMLHelper.newXMLInputFactory();
        } catch (Exception e) {
            LOGGER.error("Issue creating XMLInputFactory", e);
            throw e;
        }
    }

    private static XMLOutputFactory newXmlOutputFactory() {
        try {
            return XMLHelper.newXMLOutputFactory();
        } catch (Exception e) {
            LOGGER.error("Issue creating XMLOutputFactory", e);
            throw e;
        }
    }

    private static XMLEventFactory newXmlEventFactory() {
        try {
            return XMLHelper.newXMLEventFactory();
        } catch (Exception e) {
            LOGGER.error("Issue creating XMLEventFactory", e);
            throw e;
        }
    }
}
