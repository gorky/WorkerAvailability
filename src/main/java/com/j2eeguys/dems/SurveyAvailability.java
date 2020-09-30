/*
 * Copyright (c) 2020 This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Collection;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Main class for running the Survey Availability to read the generated XLSX
 * file and output the results.
 * 
 * @author gorky@j2eeguys.com
 */
public class SurveyAvailability implements Runnable, Closeable {

  private static final Logger LOGGER = LoggerFactory.getLogger(SurveyAvailability.class);
  
  /**
   * Spreadsheet with the worker availability.
   */
  protected File availabilityFile;
  
  /**
   * Spreadsheet with the worker files.
   */
  protected File workerFile;
  
  /**
   * Connection URL to the Database.
   */
  protected final String url;

  /**
   * Connection to the Database.
   */
  protected Connection conn;

  /**
   * Constructor for SurveyAvailability.  Uses in-memory HSQLDB.
   */
  public SurveyAvailability() {
    this("jdbc:hsqldb:mem:result;shutdown=true");
    // end <init>
  }

  /**
   * Constructor for SurveyAvailability.
   * @param url The URL for the Database.
   */
  public SurveyAvailability(final String url) {
    this.url = url;
    // end <init>
  }

  /**
   * Create the connection to the Database.
   * @param createTables if the Database Tables should be created.  Set to true if running in standalone mode.
   * 
   * @return Connection to the Database
   * @throws SQLException thrown if the Tables can not be created.
   * @throws IOException  thrown if the Init SQL Statements can not be
   *                        loaded/read.
   */
  protected Connection setupDB(final boolean createTables) throws SQLException, IOException {
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
   * @param args Command line arguments for the program.
   * @throws Exception thrown if any failures occur during execution.
   */
  public static void main(String[] args) throws Exception {
//    final File workerFile = selectSourceFile(FileSystemView.getFileSystemView().getHomeDirectory(), "Worker Spreadsheet");
    final File workerFile = selectSourceFile(new File("/tmp"), "Worker Spreadsheet");
    final File availabilityFile = selectSourceFile(workerFile == null ?
        FileSystemView.getFileSystemView().getHomeDirectory() :
          workerFile.getParentFile(),
        "Availability Spreadsheet");
    if (availabilityFile != null) {
      try (final SurveyAvailability surveyAvailability = new SurveyAvailability();){
        surveyAvailability.availabilityFile = availabilityFile;
        surveyAvailability.workerFile = workerFile;
        surveyAvailability.conn = surveyAvailability.setupDB(true);
        surveyAvailability.run();
      }
    }
    System.out.println("Finished.");
    // end main
  }

  /**
   * Select the Survey File to read.
   * @param directory Default Directory to list files in for selection.
   * @param title The title to display for the File Chooser Window.
   * @return Handle to the selected survey file.
   */
  public static File selectSourceFile(final File directory, final String title) {
    JFileChooser jfc = new JFileChooser(directory);
    FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX Files", "xlsx", "XLSX");
    jfc.setAcceptAllFileFilterUsed(false);
    jfc.addChoosableFileFilter(filter);
    jfc.setDialogTitle(title);

    int returnValue = jfc.showOpenDialog(null);

    if (returnValue == JFileChooser.APPROVE_OPTION) {
      return jfc.getSelectedFile();
    } // else, nothing selected
    return null;
    //end selectSourceFile
  }

  /*
   * (non-Javadoc)
   * @see java.lang.Runnable#run()
   */
  @Override
  public void run() {
    try {
      if (this.workerFile != null) {
        new ParseWorkerXLSX(this.workerFile, this.conn).process();
      }
      final ParseResultXLSX parseResultXLSX = new ParseResultXLSX(this.availabilityFile, this.conn, false);
      parseResultXLSX.process();
      final CellStyle headerStyle = parseResultXLSX.getHeaderStyle();
      LOGGER.info("Writing.....");
      new WriteXLSX(this.availabilityFile.getParentFile(), this.conn, headerStyle).write();
      LOGGER.info("XLSX Created.");
    } catch (IOException e) {
      throw new RuntimeException("Exception processing " + this.availabilityFile, e);
    }
    // end run
  }

  /* (non-Javadoc)
   * @see java.io.Closeable#close()
   */
  @Override
  public void close() throws IOException {
    try {
      if (this.conn != null && !this.conn.isClosed()) {
        this.conn.close();
        this.conn = null;
      }
    } catch (SQLException e) {
      throw new IOException("Exception closing SQL Connection to Database: " + this.url, e);
    }
    //end close
  }

}
