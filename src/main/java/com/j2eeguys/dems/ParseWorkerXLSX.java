/*
 * Copyright (c) 2020 This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import java.io.File;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Parses the Poll Worker Spreadsheet and loads the Database with the supplied
 * information.
 * 
 * @author gorky@j2eeguys.com
 */
public class ParseWorkerXLSX extends AbstractParserXLSX {

  /**
   * Constructor for ParserWorkerXLSX.
   * 
   * @param sourceFile The file with the Worker information.
   * @param c          Connection to the Database.
   */
  protected ParseWorkerXLSX(final File sourceFile, final Connection c) {
    super(sourceFile, c);
    // end <init>
  }

  /*
   * (non-Javadoc)
   * @see com.j2eeguys.dems.AbstractParserXLSX#load(org.apache.poi.xssf.usermodel.
   * XSSFWorkbook)
   */
  @Override
  protected void load(Workbook workbook) throws SQLException {
    try (
        final PreparedStatement search = this.c
            .prepareStatement("SELECT ID, NOTES, Email FROM WORKER WHERE LAST_NAME = ? AND FIRST_NAME = ?");
        final PreparedStatement insertWorker = this.c.prepareStatement(
            "INSERT INTO WORKER (LAST_NAME, FIRST_NAME, CITY, PHONE, EMAIL, EXPERIENCED, LANGUAGES, LOCATION, NOTES) "
                + "VALUES (?,?,?,?,?,?,?,?,?)");
        final PreparedStatement updateWorker =
            this.c.prepareStatement("UPDATE WORKER SET NOTES = ?, EMAIL = ? " + "WHERE ID = ?");

    ) {
      for (int i = 1; i <= 2; i++) {
        final Sheet currentSheet = workbook.getSheetAt(i);
        loadWorkerInfo(search, insertWorker, updateWorker, currentSheet);
      } // end for
    } // end try
      // end load
  }

  /**
   * Load the workerInfo from the currentSheet.
   * 
   * @param search       {@link PreparedStatement} for searching for an existing
   *                       record.
   * @param insertWorker {@link PreparedStatement} for inserting a new record.
   * @param updateWorker {@link PreparedStatement} for updating a worker record.
   * @param currentSheet The sheet currently being parsed.
   * @throws SQLException if any failures occur talking to the database.
   */
  protected void loadWorkerInfo(final PreparedStatement search, final PreparedStatement insertWorker,
      final PreparedStatement updateWorker, final Sheet currentSheet) throws SQLException {
    final String sheetName = currentSheet.getSheetName();
    this.LOGGER.info("Working Sheet {}", sheetName);
    final int rowCount = currentSheet.getLastRowNum();
    for (int j = 0; j < rowCount; j++) {
      final Row row = currentSheet.getRow(j);
      final Cell firstNameCell = row.getCell(1);
      if (j == 0) {
        // Header Row, let's do Sanity check
        if (!(firstNameCell.getStringCellValue().equals("First Name")
            && row.getCell(2).getStringCellValue().equals("Last Name")
            && row.getCell(3).getStringCellValue().equals("City")
            && row.getCell(4).getStringCellValue().equals("Phone #")
            && row.getCell(5).getStringCellValue().equals("Email")
            && row.getCell(6).getStringCellValue().equals("Poll Worker Exp.")
            && row.getCell(7).getStringCellValue().equals("Proficient in another language?"))) {
          StringBuilder sb = new StringBuilder(255);
          for (int m = 1; m <= 7; m++) {
            sb.append(row.getCell(m).getStringCellValue());
            sb.append(',');
          }
          sb.deleteCharAt(sb.length() - 1);
          this.LOGGER.warn("Incorrect Header Order/Missing Headers:\n{}", sb);
          // Skip to inserts.
        } else { // we're good to go. Update header information and start next row.
          this.headerStyle = row.isFormatted() ? row.getRowStyle() : firstNameCell.getCellStyle();
          continue;
        }
      } // else
      if (firstNameCell == null) {
        // Empty Row
        continue;
      }
      try {
        final String fName = firstNameCell.getStringCellValue();
        if (fName != null && !fName.isEmpty()) {
          insertWorkerInfo(search, insertWorker, updateWorker, row);
        } // else, empty row so skip.
      } catch (IllegalStateException e) {
        StringBuilder sb = new StringBuilder(255);
        for (int m = 0; m <= 6; m++) {
          sb.append(row.getCell(m).getStringCellValue());
          sb.append(',');
        }
        sb.deleteCharAt(sb.length() - 1);
        this.LOGGER.warn("Unable to insert data for: {}", sb);
        throw e;
      }

    } // end for j
      // end loadWorkerInfo
  }

  /**
   * Inserts the PollWorker Info into the Database.
   * 
   * @param search       The {@link PreparedStatement} to use to see if the Worker
   *                       is already in the DB.
   * @param insertWorker The {@link PreparedStatement} to use to insert the Worker
   *                       into the DB.
   * @param updateWorker {@link PreparedStatement} for updating a worker record.
   * @param row          The Row from the Survey sheet with the worker data.
   * @throws SQLException Thrown if the search or insert fail.
   */
  protected void insertWorkerInfo(final PreparedStatement search, final PreparedStatement insertWorker,
      final PreparedStatement updateWorker, final Row row) throws SQLException {
    final String firstName = row.getCell(1).getStringCellValue().trim();
    final String lastName = row.getCell(2).getStringCellValue().trim();
    search.setString(1, lastName);
    search.setString(2, firstName);

    try (final ResultSet searchResult = search.executeQuery()) {
      if (!searchResult.next()) {
        insertWorker.setString(1, lastName);
        insertWorker.setString(2, firstName);
        // City
        insertWorker.setString(3, row.getCell(3).getStringCellValue().trim());
        // Phone #
        insertWorker.setString(4, row.getCell(4).getStringCellValue().trim());
        // Email
        final Cell email = row.getCell(5);
        if (email != null && email.getStringCellValue().contains("@")) {
          insertWorker.setString(5, email.getStringCellValue().trim());
        } else {
          insertWorker.setString(5, null);
        }
        // Experienced
        final Cell experienced = row.getCell(6);
        if (experienced == null) {
          insertWorker.setBoolean(6, false);
        } else {
          final String strExp = experienced.getStringCellValue().trim();
          insertWorker.setBoolean(6, !strExp.isEmpty() && "Yes".equalsIgnoreCase(strExp));
        }
        // Language
        final Cell language = row.getCell(7);
        if (language != null) {
          final String strLang = language.getStringCellValue().trim();
          if (!strLang.isEmpty() && strLang.startsWith("Yes")) {
            int start = strLang.indexOf('(') + 1;
            int end = strLang.indexOf(')');
            if (start <= 0) {
              // Language not supplied
              insertWorker.setString(7, strLang);
            } else {
              insertWorker.setString(7, strLang.substring(start, end).trim());
            }
          } else {
            // Language is not "YES"
            insertWorker.setString(7, null);
          }
        } else {
          // Language is null
          insertWorker.setString(7, null);
        }
        // Location
        final Cell locationCell = row.getCell(8);
        final String location = locationCell == null ? null
            : locationCell.getCellType() == CellType.STRING ? locationCell.getStringCellValue()
                : Long.toString((long) locationCell.getNumericCellValue());
        insertWorker.setString(8, locationCell == null ? null : location);
        // Notes
        insertWorker.setString(9, row.getCell(0) == null ? null : row.getCell(0).getStringCellValue());

        try {
          if (insertWorker.executeUpdate() != 1) {
            throw new IllegalStateException("Unable to insert " + firstName + ' ' + lastName);
          }
        } catch (SQLException e) {
          this.LOGGER.error("Unable to insert {} {}", firstName, lastName);
          throw e;
        }
      } else {
        // Notes
        updateWorker.setString(1,
            row.getCell(0) == null ? searchResult.getString(2) : row.getCell(0).getStringCellValue());
        // Email
        final Cell email = row.getCell(5);
        if (email != null && email.getStringCellValue().contains("@")) {
          updateWorker.setString(2, email.getStringCellValue().trim());
        } else {
          updateWorker.setString(2, searchResult.getString(3));
        }
        // Set ID
        updateWorker.setInt(3, searchResult.getInt(1));
        if (updateWorker.executeUpdate() != 1) {
          throw new IllegalStateException("Unable to update " + firstName + ' ' + lastName);
        }

      }
    }
    // end insertWorkerInfo
  }
}// end ParseWorkerXLSX
