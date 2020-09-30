/*
 * Copyright (c) 2020 This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import java.io.File;
import java.sql.Connection;
import java.sql.Date;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Parses the Availability XLSX. Loads the data into the Database.
 * 
 * @author gorky@j2eeguys.com
 */
public class ParseResultXLSX extends AbstractParserXLSX {


  /**
   * Calendar to use for date operations.
   */
  protected final Calendar calendar = Calendar.getInstance();
  
  /**
   * Insert new Worker Info if missing.
   */
  protected final boolean insertMissing;

  /**
   * Constructor for ParseResultXLSX.
   * @param sourceFile the file being parsed.
   * @param c The Connection to the Database.
   */
  public ParseResultXLSX(final File sourceFile, final Connection c) {
    this(sourceFile, c, true);
    // end <init>
  }
  
  /**
   * Constructor for ParseResultXLSX.
   * @param sourceFile the file being parsed.
   * @param c The Connection to the Database.
   * @param insertMissing Insert a PollWorker's info if not found in the Database.
   */
  public ParseResultXLSX(final File sourceFile, final Connection c, final boolean insertMissing) {
    super(sourceFile, c);
    this.insertMissing = insertMissing;
    // end <init>
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
      final Date sheetDate, final Row row, int id) throws SQLException {
    final String yesValue = row.getCell(5).getStringCellValue();
    if (yesValue != null && yesValue.trim().equals("Checked")) {
      final String noValue = row.getCell(6).getStringCellValue();
      if (noValue != null && noValue.trim().equals("Checked")) {
        this.LOGGER.warn("Worker {} has both 'Yes' & 'No' checked for {}",
            row.getCell(2).getStringCellValue(), sheetName);
      } else {
        final String vrNum = row.getCell(2).getStringCellValue();
        insertAvailable.setInt(1, id);
        insertAvailable.setDate(2, sheetDate);
        try {
          if (insertAvailable.executeUpdate() != 1) {
            throw new IllegalStateException("Unable to insert VR " + vrNum + " for Date " + sheetDate);
          } // end insert
          this.LOGGER.debug("Inserted Availability VR# {}/{} for {}:{}", vrNum, Integer.valueOf(id), sheetDate, yesValue);
        } catch (SQLException e) {
          this.LOGGER.warn("Exception processing VR# {}/{} for {}={}:{}", vrNum, Integer.valueOf(id), sheetDate, yesValue,e.getMessage());
        }
      }
    } // end Yes Checked

  }

  /**
   * Sets the Worker VR ID.  Also, if not filtering (see {@link #insertMissing}), adds the pollworker info.
   * @param psIdentity {@link PreparedStatement} to get the DB Id for the worker.  Used for logging/debugging.
   * @param search The {@link PreparedStatement} to use to see if the Worker is already in the DB.
   * @param nameSearch The {@link PreparedStatement} to use with just First and Last names to see if the Worker is already in the DB.
   * @param insertWorker The {@link PreparedStatement} to use to insert the Worker into the DB.
   * @param updateWorker The {@link PreparedStatement} to use to insert the Worker's VR ID in the DB.
   * @param row The Row from the Survey sheet with the worker data.
   * @return the database ID of the worker. -1 if not found and not inserting.
   * @throws SQLException Thrown if the search or insert fail.
   */
  protected int setWorkerInfo(final PreparedStatement psIdentity, final PreparedStatement search,
      final PreparedStatement nameSearch, final PreparedStatement insertWorker, final PreparedStatement updateWorker, final Row row) throws SQLException {
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
        //VR_ID not set?
        nameSearch.setString(1, lastName);
        nameSearch.setString(2, firstName);
        try (final ResultSet nameSearchRS = nameSearch.executeQuery()) {
          if (!nameSearchRS.next()) {
            //Name not found.
            if (this.insertMissing) {
              return insertWorkerInfo(psIdentity, search, insertWorker, row);
            }//else, filtering instead of inserting
            this.LOGGER.debug("{} {} Not found in DB", firstName, lastName);
            return -1;
          }//else
          updateWorker.setString(1, vrId);
          //Precinct
          final Cell precinctCell = row.getCell(3);
          updateWorker.setString(2, precinctCell.getCellType() == CellType.STRING ? precinctCell.getStringCellValue()
              : Long.toString((long)precinctCell.getNumericCellValue())
                  );
          updateWorker.setString(3, row.getCell(4).getStringCellValue());
          updateWorker.setInt(4, nameSearchRS.getInt(1));
          int updated = updateWorker.executeUpdate();
          this.LOGGER.debug("Update VR# {}/Record Count: {}", vrId, Integer.valueOf(updated));
          return nameSearchRS.getInt(1);
        }//end try nameSearch
      } // else
      return searchResult.getInt(1);
    }
    //end setWorkerInfo
  }
  
  /**
   * Inserts the PollWorker Info into the Database.
   * @param psIdentity {@link PreparedStatement} to get the DB Id for the worker.  Used for logging/debugging.
   * @param search The {@link PreparedStatement} to use to see if the Worker is already in the DB.
   * @param insertWorker The {@link PreparedStatement} to use to insert the Worker into the DB.
   * @param row The Row from the Survey sheet with the worker data.
   * @return the database ID of the worker.
   * @throws SQLException Thrown if the search or insert fail.
   */
  protected int insertWorkerInfo(final PreparedStatement psIdentity, final PreparedStatement search,
      final PreparedStatement insertWorker, final Row row) throws SQLException {
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
        final Cell precint = row.getCell(3);
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
          this.LOGGER.debug("Inserted VR# {}/{}", vrId, Integer.valueOf(identity));
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
   * @throws SQLException thrown if the data can't be inserted into the Database.
   */
  @Override
  protected void load(final Workbook workbook) throws SQLException {
    final int sheetCount = workbook.getNumberOfSheets();
    try (final PreparedStatement search = this.c.prepareStatement(
            "SELECT ID, VR_ID FROM WORKER WHERE VR_ID = ? AND LAST_NAME LIKE ? AND FIRST_NAME LIKE ?");
        final PreparedStatement nameSearch = this.c.prepareStatement(
            "SELECT ID, VR_ID FROM WORKER WHERE VR_ID IS NULL AND LAST_NAME LIKE ? AND FIRST_NAME LIKE ?");
        final PreparedStatement insertWorker = this.c.prepareStatement(
            "INSERT INTO WORKER (VR_ID, LAST_NAME, FIRST_NAME, PRECINCT, ROLE) VALUES (?,?,?,?,?)");
        final PreparedStatement updateWorker =
            this.c.prepareStatement("UPDATE WORKER SET VR_ID = ?, PRECINCT = ?, ROLE = ? WHERE ID = ?");
        final PreparedStatement insertAvailable =
            this.c.prepareStatement("INSERT INTO AVAILABILITY (id, DAY) VALUES (?,?)");
        final PreparedStatement psIdentity = this.c.prepareStatement("CALL IDENTITY()");
        ) {
      for (int i = 0; i < sheetCount; i++) {
        final Sheet currentSheet = workbook.getSheetAt(i);
        final String sheetName = currentSheet.getSheetName();
        this.LOGGER.info("Working day {}", sheetName);
        // Set Month
        this.calendar.set(Calendar.MONTH, Integer.parseInt(sheetName.substring(0, 2)) - 1);
        // SetDate
        this.calendar.set(Calendar.DAY_OF_MONTH, Integer.parseInt(sheetName.substring(3, 5)));
        final Date sheetDate = new Date(this.calendar.getTimeInMillis());
        final int rowCount = currentSheet.getLastRowNum();
        for (int j = 0; j < rowCount; j++) {
          final Row row = currentSheet.getRow(j);
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
              this.LOGGER.warn("Incorrect Header Order/Missing Headers:\n{}", sb);
              //Skip to inserts.
            } else { //we're good to go.  Update header information and start next row.
              this.headerStyle = row.isFormatted() ? row.getRowStyle() : row.getCell(0).getCellStyle();
              continue;
            }
          }// else
          try {
            final int id = setWorkerInfo(psIdentity, search, nameSearch, insertWorker, updateWorker, row);
            if (id >= 0) {
              insertAvailability(insertAvailable, sheetName, sheetDate, row, id);
            } else {
              this.LOGGER.info("Skipping {} {}", row.getCell(1).getStringCellValue(),
                  row.getCell(0).getStringCellValue());
            }
          } catch (IllegalStateException e) {
            final StringBuilder sb = new StringBuilder(255);
            for(int m = 0; m<=6;m++) {
              sb.append(row.getCell(m).getStringCellValue());
              sb.append(',');
            }
            sb.deleteCharAt(sb.length() - 1);
            this.LOGGER.warn("Unable to insert data for: {}", sb);
            throw e;
          }
        } // end for j
      } // end for i
    }
  }


}
