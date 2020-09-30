/*
 * Copyright (c) 2020 This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Writes the availability data into an XLSX file.
 * 
 * @author gorky@j2eeguys.com
 */
public class WriteXLSX {

  /**
   * Logger for the class.
   */
  private final static Logger LOGGER = LoggerFactory.getLogger(WriteXLSX.class);

  /**
   * Calendar to use for date operations.
   */
  protected final Calendar calendar = Calendar.getInstance();

  /**
   * Connection to the Database.
   */
  protected final Connection c;

  /**
   * Survey file being parsed.
   */
  protected final File destinationDir;

  /**
   * Header style to be used. Loaded from the Source file.
   */
  protected CellStyle headerStyle;

  /**
   * Constructor for WriteXLSX.
   * 
   * @param destinationDir the directory to write the output file.
   * @param c              the Connection to the Database.
   * @param headerStyle    the Cell Style to user for the HeaderRow.
   */
  public WriteXLSX(final File destinationDir, final Connection c, final CellStyle headerStyle) {
    this.c = c;
    this.destinationDir = destinationDir;
    this.headerStyle = headerStyle;
    // end <init>
  }

  /**
   * Processes the survey result file.
   * 
   * @throws IOException thrown if an exception occurs during processing.
   */
  public void write() throws IOException {
    try (final Workbook outBook = buildOutput();
        final OutputStream out =
            new FileOutputStream(new File(this.destinationDir, "WorkerAvailability.xlsx"));) {
      outBook.write(out);
      out.flush();
    } catch (SQLException e) {
      throw new IOException(e.getMessage(), e);
    }
    // end process
  }

  /**
   * Builds the output sheet from the Data in the Database.
   * 
   * @return Workbook with the Workers availability.
   * @throws SQLException thrown if data can not be read from the database.
   * @throws IOException  thrown if the output workbook can not be created.
   */
  protected Workbook buildOutput() throws SQLException, IOException {
    final Workbook workbook = WorkbookFactory.create(true);
    final CellStyle cellStyle = workbook.createCellStyle();
    cellStyle.cloneStyleFrom(this.headerStyle);
    this.headerStyle = cellStyle;
    final CellStyle centerStyle = workbook.createCellStyle();
    centerStyle.setAlignment(HorizontalAlignment.CENTER);
    buildMainSheet(workbook, centerStyle);
    buildDetailSheets(workbook, centerStyle);
    buildNotScheduled(workbook, centerStyle);
    return workbook;
    // end buildOutput
  }

  /**
   * Build the tab showing the workers that have not been scheduled.
   * 
   * @param workbook    the workbook showing workers availability and information.
   * @param centerStyle Style to use for centering in the various Fields.
   * @throws SQLException thrown if any faults occur accessing the Database.
   */
  protected void buildNotScheduled(final Workbook workbook, final CellStyle centerStyle) throws SQLException {
    final Sheet sheet = workbook.createSheet("NotScheduled");
    LOGGER.info("Working sheet {}", sheet.getSheetName());
    try (final PreparedStatement listWorker = this.c.prepareStatement(
        "SELECT NOTES, LAST_NAME, FIRST_NAME, VR_ID, CITY, PHONE, EMAIL, EXPERIENCED, LANGUAGES, LOCATION, "
            + "PRECINCT, ROLE, id FROM WORKER W WHERE W.VR_ID IS NULL OR W.ID NOT IN (SELECT DISTINCT A.ID FROM AVAILABILITY A) "
            + "ORDER BY LAST_NAME, FIRST_NAME");
        final PreparedStatement searchAvailability =
            this.c.prepareStatement("SELECT DAY FROM AVAILABILITY WHERE ID = ? ORDER BY DAY");) {
      addHeaderRow(sheet, 0, 0, "Note", "Last Name", "First Name", "VR #", "City", "Phone", "Email",
          "Experienced", "Languages", "Location", "Precinct", "Role");
      addRows(centerStyle, listWorker, searchAvailability, sheet);
    }
    // end buildMainSheet
  }

  /**
   * Builds the Main Sheet with Worker Details and all available days.
   * 
   * @param workbook    The workbook to add the sheet to.
   * @param centerStyle Style to use for centering in the various Fields.
   * @throws SQLException thrown if any faults occur accessing the Database.
   */
  protected void buildMainSheet(final Workbook workbook, final CellStyle centerStyle) throws SQLException {
    final Sheet sheet = workbook.createSheet("Workers");
    LOGGER.info("Working sheet {}", sheet.getSheetName());
    try (final PreparedStatement listWorker = this.c.prepareStatement(
        "SELECT NOTES, LAST_NAME, FIRST_NAME, VR_ID, CITY, PHONE, EMAIL, EXPERIENCED, LANGUAGES, LOCATION, "
            + "PRECINCT, ROLE, id " + "FROM WORKER ORDER BY LAST_NAME, FIRST_NAME");
        final PreparedStatement searchAvailability =
            this.c.prepareStatement("SELECT DAY FROM AVAILABILITY WHERE ID = ? ORDER BY DAY");) {
      addHeaderRow(sheet, 13, 30 - 12, "Note", "Last Name", "First Name", "VR #", "City", "Phone", "Email",
          "Experienced", "Languages", "Location", "Precinct", "Role");
      addRows(centerStyle, listWorker, searchAvailability, sheet);
      // end buildMainSheet
    }

    // end buildMainSheet
  }

  /**
   * Adds the Worker Rows to the Spreadsheet.
   * 
   * @param centerStyle        Style to use for centering in the various Fields.
   * @param listWorker         query for the worker list.
   * @param searchAvailability Query to find the Availability for a given worker.
   * @param sheet              the SHeet to add the Rows to.
   * @throws SQLException thrown if any faults occur accessing the Database.
   */
  protected void addRows(final CellStyle centerStyle, final PreparedStatement listWorker,
      final PreparedStatement searchAvailability, final Sheet sheet) throws SQLException {
    try (final ResultSet rsWorker = listWorker.executeQuery()) {
      final ResultSetMetaData listMetaData = listWorker.getMetaData();
      // NOTE: id column is the last in the list.
      final int colCount = listMetaData.getColumnCount();
      final int offSet = 13 - colCount + 1;
      int rowNum = 1;
      while (rsWorker.next()) {
        // Load Names
        final Row workerRow = sheet.createRow(rowNum++);
        for (int k = 0, l = 1; l < colCount; k++, l++) {
          final Cell cell = workerRow.createCell(k);
          LOGGER.debug("Cell Type: {}", listMetaData.getColumnTypeName(l));
          if (listMetaData.getColumnType(l) == Types.BOOLEAN) {
            if (rsWorker.getBoolean(l)) {
              cell.setCellValue("X");
              cell.setCellStyle(centerStyle);
            }
          } else if (listMetaData.getColumnType(l) == Types.CHAR) {
            if (rsWorker.getByte(l) > 0) {
              cell.setCellValue("Yes");
              cell.setCellStyle(centerStyle);
            }
          } else if (listMetaData.getColumnType(l) == Types.SMALLINT) {
            if (rsWorker.getInt(l) > 0) {
              cell.setCellValue(Integer.toString(rsWorker.getInt(l)));
            }
          } else {
            cell.setCellValue(rsWorker.getString(l));
          }
        }
        // Load Availability
        LOGGER.debug("Loading {}/{}", rsWorker.getString(4), Integer.valueOf(rsWorker.getInt(colCount)));
        searchAvailability.setInt(1, rsWorker.getInt(colCount));
        try (final ResultSet rsAvailable = searchAvailability.executeQuery();) {
          while (rsAvailable.next()) {
            final Date day = rsAvailable.getDate(1);
            LOGGER.debug("Available: {}", day);
            this.calendar.setTimeInMillis(day.getTime());
            int dayInMonth = this.calendar.get(Calendar.DAY_OF_MONTH);
            final Cell cell = workerRow.createCell(dayInMonth - offSet);
            // cell.setCellValue(dayInMonth);
            cell.setCellValue("X");
            cell.setCellStyle(centerStyle);
          }
        }
      } // end rsWorker
    }
    // end addRow
  }

  /**
   * Builds the output sheet from the Data in the Database.
   * 
   * @param workbook    The workbook to add the Availability Details to.
   * @param centerStyle Style to use for centering in the various Fields.
   * @throws SQLException thrown if data can not be read from the database.
   * @throws IOException  thrown if the output workbook can not be created.
   */
  protected void buildDetailSheets(final Workbook workbook, final CellStyle centerStyle)
      throws SQLException, IOException {
    final CellStyle cellStyle = workbook.createCellStyle();
    cellStyle.cloneStyleFrom(this.headerStyle);
    this.headerStyle = cellStyle;
    centerStyle.setAlignment(HorizontalAlignment.CENTER);
    try (
        final PreparedStatement searchAvailability = this.c.prepareStatement(
            "SELECT DAY FROM AVAILABILITY WHERE ID = ? AND DAY >= ? AND DAY < ? ORDER BY DAY");
        final PreparedStatement listWorker = this.c.prepareStatement(
            "SELECT LAST_NAME, FIRST_NAME, VR_ID, PRECINCT, ROLE, id FROM WORKER ORDER BY LAST_NAME, FIRST_NAME");) {
      for (int i = 12; i < 30; i += 7) {
        final Sheet sheet = workbook.createSheet("Oct " + i + '-' + (i >= 23 ? 30 : i + 7));
        LOGGER.info("Working sheet {}", sheet.getSheetName());
        addHeaderRow(sheet, i, 7, "Last Name", "First Name", "VR #", "Precinct", "Role");
        this.calendar.set(Calendar.DAY_OF_MONTH, i);
        final Date queryStartDate = new Date(this.calendar.getTimeInMillis());
        searchAvailability.setDate(2, queryStartDate);
        this.calendar.set(Calendar.DAY_OF_MONTH, (i >= 23 ? 31 : i + 7));
        final Date queryEndDate = new Date(this.calendar.getTimeInMillis());
        searchAvailability.setDate(3, queryEndDate);
        try (final ResultSet rsWorker = listWorker.executeQuery()) {
          int rowNum = 1;
          while (rsWorker.next()) {
            // Load Names
            final Row workerRow = sheet.createRow(rowNum++);
            for (int k = 0, l = 1; k < 5; k++, l++) {
              final Cell cell = workerRow.createCell(k);
              cell.setCellValue(rsWorker.getString(l));
            }
            // Load Availability
            LOGGER.debug("Loading {}/{}", rsWorker.getString(3), Integer.valueOf(rsWorker.getInt(6)));
            searchAvailability.setInt(1, rsWorker.getInt(6));
            try (final ResultSet rsAvailable = searchAvailability.executeQuery();) {
              while (rsAvailable.next()) {
                final Date day = rsAvailable.getDate(1);
                LOGGER.debug("Available: {}", day);
                this.calendar.setTimeInMillis(day.getTime());
                int dayInMonth = this.calendar.get(Calendar.DAY_OF_MONTH);
                final Cell cell = workerRow.createCell(dayInMonth - i + 5);
                // cell.setCellValue(dayInMonth);
                cell.setCellValue("X");
                cell.setCellStyle(centerStyle);
              }
            }
          } // end rsWorker
        }
      }
    }
    // end buildDetailSheets
  }

  /**
   * Add a header row to the Spreadsheet.
   * 
   * @param sheet  The Sheet to add the Header row to.
   * @param start  the Starting Date for the Sheet.
   * @param cols   The number of date columns on the Spreadsheet.
   * @param titles The titles for the Header Row
   */
  protected void addHeaderRow(final Sheet sheet, final int start, final int cols, final String... titles) {
    final Row headerRow = sheet.createRow(0);
    this.cellNum = 0;
    for (final String title : titles) { createHeaderCell(headerRow, title); }
    for (int i = 0; i < cols && (start + i) <= 30; i++) {
      createHeaderCell(headerRow, Integer.toString(start + i));
    }
    // end addHeaderRow
  }

  private int cellNum;

  /**
   * Create a header Cell in the row for the columns in the sheet.
   * 
   * @param headerRow The row that the headers/titles are being added to.
   * @param cellValue The title for the column.
   */
  protected void createHeaderCell(final Row headerRow, final String cellValue) {
    Cell cell = headerRow.createCell(this.cellNum++);
    cell.setCellStyle(this.headerStyle);
    cell.setCellValue(cellValue);
    // end createHeaderCell
  }
}
