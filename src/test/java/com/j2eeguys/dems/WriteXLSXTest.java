/*
 * Copyright (c) 2020
 * 
 * This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory;
import org.apache.poi.ss.usermodel.Row;
import org.junit.jupiter.api.Test;

/**
 * @author gorky@j2eeguys.com
 *
 */
class WriteXLSXTest {

  /**
   * Test method for {@link com.j2eeguys.dems.WriteXLSX#addHeaderRow(org.apache.poi.ss.usermodel.Sheet, int, int, String...)}.
   * @throws IOException thrown if an exception occurs during testing.
   */
  @Test
  void testAddHeader() throws IOException {
    final WriteXLSX parser = new WriteXLSX(new File("build.gradle"), null, null);
    try (final HSSFWorkbook workbook = HSSFWorkbookFactory.createWorkbook();
        ){
      final HSSFSheet sheet = workbook.createSheet("Oct 10");
      parser.addHeaderRow(sheet, 12, 7, "Last Name", "First Name", "VR #", "Precinct", "Role" );
      assertEquals(0, sheet.getLastRowNum());
      assertEquals(1, sheet.getPhysicalNumberOfRows());
      Row row = sheet.getRow(0);
      assertEquals(12, row.getPhysicalNumberOfCells());
    }
    //end testAddHeader
  }

}
