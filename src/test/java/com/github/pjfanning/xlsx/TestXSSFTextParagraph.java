/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package com.github.pjfanning.xlsx;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.awt.*;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import static org.junit.Assert.*;

public class TestXSSFTextParagraph {
    @Test
    public void testXSSFTextParagraph() throws IOException {
        try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
            try (XSSFWorkbook wb = new XSSFWorkbook()) {
                XSSFSheet sheet = wb.createSheet();
                XSSFDrawing drawing = sheet.createDrawingPatriarch();

                XSSFTextBox shape = drawing.createTextbox(new XSSFClientAnchor(0, 0, 0, 0, 2, 2, 3, 4));
                XSSFRichTextString rt = new XSSFRichTextString("Test String");

                XSSFFont font = wb.createFont();
                Color color = new Color(0, 255, 255);
                font.setColor(new XSSFColor(color, wb.getStylesSource().getIndexedColors()));
                font.setFontName("Arial");
                rt.applyFont(font);

                shape.setText(rt);

                wb.write(bos);
            }


            try (Workbook workbook = StreamingReader.builder().setReadShapes(true)
                            .open(bos.toInputStream())) {
                Drawing<?> drawing = workbook.getSheetAt(0).getDrawingPatriarch();
                assertNotNull(drawing);
                Iterator<?> iterator = drawing.iterator();
                assertTrue("drawing iterator has a shape?", iterator.hasNext());
                XSSFSimpleShape shape = (XSSFSimpleShape) drawing.iterator().next();

                List<XSSFTextParagraph> paras = shape.getTextParagraphs();
                assertEquals(1, paras.size());

                XSSFTextParagraph text = paras.get(0);
                assertEquals("Test String", text.getText());
                assertFalse(text.isBullet());

                assertNotNull(text.getTextRuns());
                assertEquals(1, text.getTextRuns().size());

                assertEquals(TextAlign.LEFT, text.getTextAlign());
            }
        }
    }
}
