package com.github.pjfanning.xlsx.impl.constants;

import javax.xml.namespace.QName;

public enum QNameBucket {
	QNAME_HIDDEN("hidden"),
	QNAME_HT("ht"),
	QNAME_MAX("max"),
	QNAME_MIN("min"),
	QNAME_R("r"),
	QNAME_REF("ref"),
	QNAME_S("s"),
	QNAME_T("t"),
	QNAME_WIDTH("width"),
	CONFORMANCE("conformance");

	private final QName qName;

	QNameBucket(String localPart) {
		qName = QName.valueOf(localPart);
	}

	public QName getQName() {
		return qName;
	}
}
