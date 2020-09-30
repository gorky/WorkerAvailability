/*
 * Copyright (c) 2020
 * 
 * This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import static org.junit.jupiter.api.Assertions.assertFalse;

import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.junit.jupiter.api.Test;

/**
 * @author gorky@j2eeguys.com
 *
 */
class SurveyAvailabilityTest {

  /**
   * Test method for {@link com.j2eeguys.dems.SurveyAvailability#setupDB(boolean)}.
   * @throws IOException 
   * @throws SQLException 
   */
  @Test
  void testSetupDB() throws SQLException, IOException {
    try (final SurveyAvailability sa = new SurveyAvailability();
        final Connection c = sa.setupDB(true);
        final Statement s = c.createStatement();){
      try (final ResultSet rs = s.executeQuery(("SELECT * FROM WORKER"));){
        assertFalse(rs.next(), "Should be empty table");
      }
    }
    //end testSetupDB
  }

}
