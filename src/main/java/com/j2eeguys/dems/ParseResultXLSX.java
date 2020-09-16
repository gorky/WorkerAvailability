/*
 * Copyright (c) 2020 This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Calendar;
import java.util.Collection;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Parses the input XLSX. Loads the data into the Database.
 * 
 * @author gorky@j2eeguys.com
 */
public class ParseResultXLSX {

  private static final Logger LOGGER = LoggerFactory.getLogger(ParseResultXLSX.class);

  /**
   * Calendar to use for date operations.
   */
  protected final Calendar calendar = Calendar.getInstance();

  /**
   * Survey file being parsed.
   */
  protected final File sourceFile;
  
  /**
   * Connection URL to the Database.
   */
  protected final String url;

  /**
   * Header style to be used.  Loaded from the Source Surbery file.
   */
  protected CellStyle headerStyle;

  /**
   * Constructor for ParseResultXLSX.
   * @param sourceFile The source survey file to read.
   */
  public ParseResultXLSX(final File sourceFile) {
    if (!sourceFile.exists() && !sourceFile.canRead()) {
      throw new IllegalArgumentException("Unable to Read " + sourceFile.getAbsolutePath());
    }
    this.sourceFile = sourceFile;
    this.url = "jdbc:hsqldb:mem:result;shutdown=true";
    // end <init>
  }

  /**
   * Processes the survey result file.
   * 
   * @throws IOException thrown if an exception occurs during processing.
   */
  public void process() throws IOException {
    try (final XSSFWorkbook workbook = XSSFWorkbookFactory.createWorkbook(this.sourceFile, true);
        // Prep table
        final Connection c = prepTable();) {
      load(workbook, c);
      //Save Processed Data!
      try (final Workbook outBook = buildOutput(c);
          final OutputStream out =
          new FileOutputStream(new File(this.sourceFile.getParentFile(), "WorkerAvailability.xlsx"));) {
        outBook.write(out);
        out.flush();
      }

    } catch (SQLException e) {
      throw new IOException(e.getMessage(), e);
    }
    // end process
  }
  
  /**
   * Insert that a worker is available for a given date.
   * @param insertAvailable {@link PreparedStatement} for inserting the worker's availability info.
   * @param sheetName Name of the sheet being handled (for logging purposes).
   * @param sheetDate Date being handled.
   * @param row the Row with the details for the worker.
   * @param id Database ID for the Worker.
   * @throws SQLException thrown if the availability information can not be added to the Database.
   */
  protected void insertAvailability(final PreparedStatement insertAvailable, final String sheetName,
      final Date sheetDate, final XSSFRow row, int id) throws SQLException {
    final String yesValue = row.getCell(5).getStringCellValue();
    if (yesValue != null && yesValue.trim().equals("Checked")) {
      final String noValue = row.getCell(6).getStringCellValue();
      if (noValue != null && noValue.trim().equals("Checked")) {
        LOGGER.warn("Worker {} has both 'Yes' & 'No' checked for {}",
            row.getCell(2).getStringCellValue(), sheetName);
      } else {
        final String vrNum = row.getCell(2).getStringCellValue();
        insertAvailable.setInt(1, id);
        insertAvailable.setDate(2, sheetDate);
        if (insertAvailable.executeUpdate() != 1) {
          throw new IllegalStateException("Unable to insert VR " + vrNum + " for Date " + sheetDate);
        } // end insert
        LOGGER.debug("Inserted Availability VR# {}/{} for {}:{}", vrNum, Integer.valueOf(id), sheetDate, yesValue);
      }
    } // end Yes Checked

  }

  /**
   * Inserts the PollWorker Info into the Database.
   * @param psIdentity 
   * @param search The {@link PreparedStatement} to use to see if the Worker is already in the DB.
   * @param insertWorker The {@link PreparedStatement} to use to insert the Worker into the DB.
   * @param row The Row from the Survey sheet with the worker data.
   * @return the database ID of the worker.
   * @throws SQLException Thrown if the search or insert fail.
   */
  protected int insertWorkerInfo(final PreparedStatement psIdentity, final PreparedStatement search,
      final PreparedStatement insertWorker, final XSSFRow row) throws SQLException {
    final String vrId = row.getCell(2).getStringCellValue().trim();
    search.setString(1, vrId);
    final String lastName = row.getCell(0).getStringCellValue().trim();
    final String firstName = row.getCell(1).getStringCellValue().trim();
    if (vrId.trim().length() == 0 || !Character.isDigit(vrId.charAt(0))) {
      //Empty or non-numeric
      search.setString(2, lastName);
      search.setString(3, firstName);
    } else {
      //VR # was numeric, and that's unique per person.
      search.setString(2, "%");
      search.setString(3, "%");
    }
    try (final ResultSet searchResult = search.executeQuery()) {
      if (!searchResult.next()) {
        insertWorker.setObject(1, vrId);
        insertWorker.setString(2, lastName);
        insertWorker.setString(3, firstName);
        final XSSFCell precint = row.getCell(3);
        if (precint.getCellType() == CellType.NUMERIC) {
          insertWorker.setInt(4, (int) precint.getNumericCellValue());
        } else {
          final String value = precint.getStringCellValue().trim();
          if (value.length() > 0 ) {
            insertWorker.setInt(4, Integer.parseInt(value));
          } else {
            insertWorker.setObject(4, null);
          }
        }
        insertWorker.setString(5, row.getCell(4).getStringCellValue());
        if (insertWorker.executeUpdate() != 1) {
          throw new IllegalStateException("Unable to insert VR " + vrId);
        }//else
        try (final ResultSet rsId = psIdentity.executeQuery();){
          rsId.next();
          int identity = rsId.getInt(1);
          LOGGER.debug("Inserted VR# {}/{}", vrId, Integer.valueOf(identity));
          return identity;
        }
      }//else
      return searchResult.getInt(1);
    }
    // end insertWorkerInfo
  }

  /**
   * Load the Survey data from a spreadsheet into the Database.
   * @param workbook The workbook supplying the worker availability data.
   * @param c the {@link Connection} to the database.
   * @throws SQLException thrown if the data can't be inserted into the Database.
   */
  protected void load(final XSSFWorkbook workbook, final Connection c) throws SQLException {
    final int sheetCount = workbook.getNumberOfSheets();
    try (final PreparedStatement search = c.prepareStatement(
            "SELECT ID FROM WORKER WHERE VR_ID = ? AND LAST_NAME LIKE ? AND FIRST_NAME LIKE ?");
        final PreparedStatement insertWorker = c.prepareStatement(
            "INSERT INTO WORKER (VR_ID, LAST_NAME, FIRST_NAME, PRECINT, ROLE) VALUES (?,?,?,?,?)");
        final PreparedStatement insertAvailable =
            c.prepareStatement("INSERT INTO AVAILABILITY (id, DAY) VALUES (?,?)");
        final PreparedStatement psIdentity = c.prepareStatement("CALL IDENTITY()");
        ) {
      for (int i = 0; i < sheetCount; i++) {
        final XSSFSheet currentSheet = workbook.getSheetAt(i);
        final String sheetName = currentSheet.getSheetName();
        LOGGER.info("Working day {}", sheetName);
        // Set Month
        this.calendar.set(Calendar.MONTH, Integer.parseInt(sheetName.substring(0, 2)) - 1);
        // SetDate
        this.calendar.set(Calendar.DAY_OF_MONTH, Integer.parseInt(sheetName.substring(3, 5)));
        final Date sheetDate = new Date(this.calendar.getTimeInMillis());
        final int rowCount = currentSheet.getLastRowNum();
        for (int j = 0; j < rowCount; j++) {
          final XSSFRow row = currentSheet.getRow(j);
          if (j == 0) {
            // Header Row, let's do Sanity check
            if (!(row.getCell(0).getStringCellValue().equals("Last Name")
                && row.getCell(1).getStringCellValue().equals("First Name")
                && row.getCell(2).getStringCellValue().equals("VR #")
                && row.getCell(3).getStringCellValue().equals("Precinct")
                && row.getCell(4).getStringCellValue().equals("Role")
                && row.getCell(5).getStringCellValue().equals("Yes")
                && row.getCell(6).getStringCellValue().equals("No"))) {
              StringBuilder sb = new StringBuilder(255);
              for(int m = 0; m<=6;m++) {
                sb.append(row.getCell(m).getStringCellValue());
                sb.append(',');
              }
              sb.deleteCharAt(sb.length() - 1);
              LOGGER.warn("Incorrect Header Order/Missing Headers:\n{}", sb);
              //Skip to inserts.
            } else { //we're good to go.  Update header information and start next row.
              this.headerStyle = row.isFormatted() ? row.getRowStyle() : row.getCell(0).getCellStyle();
              continue;
            }
          }// else
          try {
            final int id = insertWorkerInfo(psIdentity, search, insertWorker, row);
            insertAvailability(insertAvailable, sheetName, sheetDate, row, id);
          } catch (IllegalStateException e) {
            StringBuilder sb = new StringBuilder(255);
            for(int m = 0; m<=6;m++) {
              sb.append(row.getCell(m).getStringCellValue());
              sb.append(',');
            }
            sb.deleteCharAt(sb.length() - 1);
            LOGGER.warn("Unable to insert data for: {}", sb);
            throw e;
          }
        } // end for j
      } // end for i
    }
  }

  /**
   * Create the tables to use.
   * 
   * @return Connection to the Database
   * @throws SQLException thrown if the Tables can not be created.
   * @throws IOException  thrown if the Init SQL Statements can not be
   *                        loaded/read.
   */
  protected Connection prepTable() throws SQLException, IOException {
    final Connection c = DriverManager.getConnection(this.url, "SA", "");
    String sql = "";
    try (final Statement s = c.createStatement();
        final InputStream initSql = getClass().getResourceAsStream("/com/j2eeguys/dems/hsqldb/InitDB.sql");) {
      final Collection<String> lines = IOUtils.readLines(
          initSql, Charset.defaultCharset());
      for (final String line : lines) {
        if (line.trim().length() == 0) {
          // empty line, skip
          continue;
        } // else
        sql += line + '\n';
        if (line.endsWith(";")) {
          LOGGER.info("Executing SQL --> {}", sql);
          s.execute(sql);
          sql = "";
          LOGGER.info("Executed");
        } // end if
      } // end for
    } // end try

    return c;
    // end prepTable
  }

  /**
   * Builds the output sheet from the Data in the Database.
   * @param c the Connection to the Database.
   * @return Workbook with the Workers availability.
   * @throws SQLException thrown if data can not be read from the database.
   * @throws IOException thrown if the output workbook can not be created.
   */
  protected Workbook buildOutput(final Connection c) throws SQLException, IOException {
    final Workbook workbook = WorkbookFactory.create(true);
    final CellStyle cellStyle = workbook.createCellStyle();
    cellStyle.cloneStyleFrom(this.headerStyle);
    this.headerStyle = cellStyle;
    final CellStyle centerStyle = workbook.createCellStyle();
    centerStyle.setAlignment(HorizontalAlignment.CENTER);
    try (final PreparedStatement searchAvailability =
            c.prepareStatement("SELECT DAY FROM AVAILABILITY WHERE ID = ? AND DAY >= ? AND DAY < ? ORDER BY DAY");
        final PreparedStatement listWorker = c.prepareStatement(
            "SELECT LAST_NAME, FIRST_NAME, VR_ID, PRECINT, ROLE, id FROM WORKER ORDER BY LAST_NAME, FIRST_NAME");
        ) {
      for (int i = 12; i < 30; i += 7) {
        final Sheet sheet = workbook.createSheet("Oct " + i + '-' + (i >= 23 ? 30 : i+7));
        LOGGER.info("Working sheet {}", sheet.getSheetName());
        addHeader(sheet, i);
        this.calendar.set(Calendar.DAY_OF_MONTH, i);
        final Date queryStartDate = new Date(this.calendar.getTimeInMillis());
        searchAvailability.setDate(2, queryStartDate);
        this.calendar.set(Calendar.DAY_OF_MONTH, (i >= 23 ? 31 : i+7));
        final Date queryEndDate = new Date(this.calendar.getTimeInMillis());
        searchAvailability.setDate(3, queryEndDate);
        try (final ResultSet rsWorker = listWorker.executeQuery()) {
          int rowNum = 1;
          while (rsWorker.next()) {
            //Load Names
            final Row workerRow = sheet.createRow(rowNum++);
            for(int k = 0, l=1; k < 5;k++, l++) {
              final Cell cell = workerRow.createCell(k);              
              cell.setCellValue(rsWorker.getString(l));
            }
            //Load Availability
            LOGGER.debug("Loading {}/{}", rsWorker.getString(3), Integer.valueOf(rsWorker.getInt(6)));
            searchAvailability.setInt(1, rsWorker.getInt(6));
            try (final ResultSet rsAvailable = searchAvailability.executeQuery();){
              while(rsAvailable.next()) {
                final Date day = rsAvailable.getDate(1);
                LOGGER.debug("Available: {}", day);
                this.calendar.setTimeInMillis(day.getTime());
                int dayInMonth = this.calendar.get(Calendar.DAY_OF_MONTH);
                final Cell cell = workerRow.createCell(dayInMonth - i + 5);
                //cell.setCellValue(dayInMonth);
                cell.setCellValue("X");
                cell.setCellStyle(centerStyle);
              }
            }
          }//end rsWorker
        }
      }
    }
    return workbook;
  }

  /**
   * Add a header row to the Spreadsheet.
   * @param sheet The Sheet to add the Header row to.
   * @param start the Starting Date for the Sheet.
   */
  protected void addHeader(final Sheet sheet, final int start) {
    final Row headerRow = sheet.createRow(0);
    final String[] titles = { "Last Name", "First Name", "VR #", "Precinct", "Role" };
    this.cellNum = 0;
    for (final String title : titles) { createHeaderCell(headerRow, title); }
    for(int i = 0; i < 7 && (start + i) <= 30; i++) {
      createHeaderCell(headerRow, Integer.toString(start + i));
    }
  }

  private int cellNum;
  
  /**
   * Create a header Cell in the row for the columns in the sheet.
   * @param headerRow The row that the headers/titles are being added to.
   * @param cellValue The title for the column.
   */
  protected void createHeaderCell(final Row headerRow, final String cellValue) {
    Cell cell = headerRow.createCell(this.cellNum++);
    cell.setCellStyle(this.headerStyle);
    cell.setCellValue(cellValue);
    //end createHeaderCell
  }

}
