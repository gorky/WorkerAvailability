/*
 * Copyright (c) 2020
 * 
 * This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory;
import org.apache.poi.ss.usermodel.Row;
import org.junit.jupiter.api.Test;

/**
 * @author gorky@j2eeguys.com
 *
 */
class ParseResultXLSXTest {

  /**
   * Test method for {@link com.j2eeguys.dems.ParseResultXLSX#prepTable()}.
   * @throws IOException 
   * @throws SQLException 
   */
  @Test
  void testPrepTable() throws SQLException, IOException {
    final ParseResultXLSX parser = new ParseResultXLSX(new File("build.gradle"));
    try (final Connection c = parser.prepTable();
        final Statement s = c.createStatement();){
      try (final ResultSet rs = s.executeQuery(("SELECT * FROM WORKER"));){
        assertFalse(rs.next(), "Should be empty table");
      }
    }
    //end testPrepTable
  }

  /**
   * Test method for {@link com.j2eeguys.dems.ParseResultXLSX#addHeader(org.apache.poi.ss.usermodel.Sheet, int)}.
   * @throws IOException thrown if an exception occurs during testing.
   */
  @Test
  void testAddHeader() throws IOException {
    final ParseResultXLSX parser = new ParseResultXLSX(new File("build.gradle"));
    try (final HSSFWorkbook workbook = HSSFWorkbookFactory.createWorkbook();
        ){
      final HSSFSheet sheet = workbook.createSheet("Oct 10");
      parser.addHeader(sheet, 10);
      assertEquals(0, sheet.getLastRowNum());
      assertEquals(1, sheet.getPhysicalNumberOfRows());
      Row row = sheet.getRow(0);
      assertEquals(12, row.getPhysicalNumberOfCells());
    }
    //end testAddHeader
  }

}
