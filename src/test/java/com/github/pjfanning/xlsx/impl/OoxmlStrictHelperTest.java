package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.impl.ooxml.OoxmlStrictHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.model.ThemesTable;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.assertNotNull;

public class OoxmlStrictHelperTest {
  @Test
  public void testThemes() throws Exception {
    try(OPCPackage pkg = OPCPackage.open(new File("src/test/resources/sample.strict.xlsx"), PackageAccess.READ)) {
      ThemesTable themes = OoxmlStrictHelper.getThemesTable(pkg);
      assertNotNull(themes.getThemeColor(ThemesTable.ThemeElement.DK1.idx));
    }
  }
}
