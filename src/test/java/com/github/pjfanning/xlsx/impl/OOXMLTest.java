package com.github.pjfanning.xlsx.impl;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class OOXMLTest {
  @Test
  public void testStrictOOXML() throws IOException, InvalidFormatException {
    try (OPCPackage pkg = OPCPackage.open(new File("src/test/resources/sample.strict.xlsx"))) {
      System.out.println("testStrictOOXML parts " + pkg.getParts().size());
      for(PackagePart part : pkg.getParts()) {
        System.out.println("part " + part.getPartName() + " content-type=" + part.getContentType());
//        for(PackageRelationship pr : part.getRelationships()) {
//          System.out.println("relationship " + pr.getRelationshipType() + " id=" + pr.getId() + " source-uri=" + pr.getSourceURI());
//        }
      }
    }
  }
}
