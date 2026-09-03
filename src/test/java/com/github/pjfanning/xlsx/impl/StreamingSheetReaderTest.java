package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.exceptions.CloseException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Assert;
import org.junit.Test;
import org.mockito.Mockito;

import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.util.List;

public class StreamingSheetReaderTest {
  @Test
  public void testStrictDates() throws Exception {
    PackagePart packagePart = Mockito.mock(PackagePart.class);
    Mockito.when(packagePart.getInputStream()).thenAnswer(I ->
            new FileInputStream("src/test/resources/strict.dates.xml"));
    StreamingSheetReader reader = new StreamingSheetReader(
            null, packagePart, null, null, null, true, 100);
    try {
      Assert.assertEquals(0, reader.getFirstRowNum());
      Assert.assertEquals(0, reader.getLastRowNum());
      Row firstRow = reader.iterator().next();
      Assert.assertEquals(CellType.NUMERIC, firstRow.getCell(0).getCellType());
      Assert.assertEquals("2021-02-28", firstRow.getCell(0).getStringCellValue());
      Assert.assertEquals(LocalDate.parse("2021-02-28").atStartOfDay(),
              firstRow.getCell(0).getLocalDateTimeCellValue());
      Assert.assertEquals(java.sql.Date.valueOf(LocalDate.parse("2021-02-28")),
              firstRow.getCell(0).getDateCellValue());
      Assert.assertEquals(CellType.NUMERIC, firstRow.getCell(1).getCellType());
      Assert.assertEquals("12:00:00.000", firstRow.getCell(1).getStringCellValue());
      Assert.assertEquals(DateUtil.convertTime("12:00"), firstRow.getCell(1).getNumericCellValue(), 0.001);
    } finally {
      reader.close();
    }
  }

  @Test
  public void testCloseClosesEveryIteratorEvenIfOneFails() throws Exception {
    PackagePart packagePart = Mockito.mock(PackagePart.class);
    Mockito.when(packagePart.getInputStream()).thenAnswer(I ->
            new FileInputStream("src/test/resources/strict.dates.xml"));
    StreamingSheetReader reader = new StreamingSheetReader(
            null, packagePart, null, null, null, true, 100);

    StreamingRowIterator failing = Mockito.mock(StreamingRowIterator.class);
    Mockito.doThrow(new CloseException("boom")).when(failing).close(false);
    StreamingRowIterator healthy = Mockito.mock(StreamingRowIterator.class);

    Field iteratorsField = StreamingSheetReader.class.getDeclaredField("iterators");
    iteratorsField.setAccessible(true);
    @SuppressWarnings("unchecked")
    List<StreamingRowIterator> iterators = (List<StreamingRowIterator>) iteratorsField.get(reader);
    iterators.clear();
    iterators.add(failing);
    iterators.add(healthy);

    //the failure is still reported...
    Assert.assertThrows(CloseException.class, reader::close);
    //...but the iterator after it was closed anyway, rather than being leaked
    Mockito.verify(healthy).close(false);
  }
}
