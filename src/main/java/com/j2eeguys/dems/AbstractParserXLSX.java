/*
 * Copyright (c) 2020
 * 
 * This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author gorky@j2eeguys.com
 *
 */
public abstract class AbstractParserXLSX {

  /**
   * Logger for the class.
   */
  protected final Logger LOGGER = LoggerFactory.getLogger(getClass());
  
  /**
   * Survey file being parsed.
   */
  protected final File sourceFile;

  /**
   * Connection to the Database.
   */
  protected final Connection c;
  
  /**
   * Header style to be used.  Loaded from the Source file.
   */
  protected CellStyle headerStyle;

  /**
   * Constructor for AbstractParserXLSX.
   * @param sourceFile the file being parsed.
   * @param c The Connection to the Database.
   */
  protected AbstractParserXLSX(final File sourceFile, final Connection c) {
    super();
    if (!sourceFile.exists() && !sourceFile.canRead()) {
      throw new IllegalArgumentException("Unable to Read " + sourceFile.getAbsolutePath());
    }
    this.sourceFile = sourceFile;
    this.c = c;
    //end <init>
  }
  
  /**
   * Load the Survey data from a spreadsheet into the Database.
   * @param workbook The workbook supplying the worker availability data.
   * @throws SQLException thrown if the data can't be inserted into the Database.
   */
  protected abstract void load(final Workbook workbook) throws SQLException;
  
  /**
   * Processes the XSLX file.
   * 
   * @throws IOException thrown if an exception occurs during processing.
   */
  public void process() throws IOException{
    try (final XSSFWorkbook workbook = XSSFWorkbookFactory.createWorkbook(this.sourceFile, true);){
      load(workbook);
    } catch (SQLException e) {
      throw new IOException("Exception processing " + this.sourceFile.getAbsolutePath(), e);
    }
    //end process
  }
  
  /**
   * @return the headerStyle
   */
  public CellStyle getHeaderStyle() {
    return this.headerStyle;
  }
  
}//end AbstractParserXLSX
