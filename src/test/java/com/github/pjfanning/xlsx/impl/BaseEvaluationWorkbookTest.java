package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.ss.formula.SheetIdentifier;
import org.apache.poi.ss.formula.ptg.NameXPxg;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

public class BaseEvaluationWorkbookTest {

  @Test
  public void testResolveBracketedBookIndex() throws Exception {
    //the [] wrapper is stripped before the index is parsed
    assertEquals(1, resolveBookIndex("[1]"));
    assertEquals(23, resolveBookIndex("[23]"));
  }

  @Test
  public void testResolveUnbracketedBookIndex() throws Exception {
    assertEquals(1, resolveBookIndex("1"));
  }

  private int resolveBookIndex(final String bookName) throws Exception {
    try (Workbook wb = StreamingReader.builder().open(new File("src/test/resources/gaps.xlsx"))) {
      CurrentRowEvaluationWorkbook evaluationWorkbook = new CurrentRowEvaluationWorkbook(wb, null);
      //a SheetIdentifier with no sheet part resolves to a workbook + named range reference
      NameXPxg ptg = (NameXPxg) evaluationWorkbook.getNameXPtg(
              "myNamedRange", new SheetIdentifier(bookName, null));
      assertNotNull(ptg);
      return ptg.getExternalWorkbookNumber();
    }
  }
}
