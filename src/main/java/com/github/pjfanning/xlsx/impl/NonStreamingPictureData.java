package com.github.pjfanning.xlsx.impl;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.usermodel.XSSFPictureData;

public class NonStreamingPictureData extends XSSFPictureData {
  public NonStreamingPictureData(PackagePart part) {
    super(part);
  }
}
