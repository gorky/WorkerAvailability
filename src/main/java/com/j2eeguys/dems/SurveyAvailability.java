/*
 * Copyright (c) 2020 This code is licensed under the GPLv2.
 */
package com.j2eeguys.dems;

import java.io.File;
import java.io.IOException;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

/**
 * Main class for running the Survey Availability to read the generated XLSX
 * file and output the results.
 * 
 * @author gorky@j2eeguys.com
 */
public class SurveyAvailability implements Runnable {

  /**
   * File being read.
   */
  protected File sourceFile;
  /**
   * Constructor for SurveyAvailability.
   */
  public SurveyAvailability() {
    // end <init>
  }

  /**
   * @param args
   */
  public static void main(String[] args) {
    final File sourceFile = selectSourceFile();
    if (sourceFile != null) {
      final SurveyAvailability surveyAvailability = new SurveyAvailability();
      surveyAvailability.sourceFile = sourceFile;
      surveyAvailability.run();
    }
    // end main
  }

  /**
   * Select the Survey File to read.
   * @return Handle to the selected survey file.
   */
  public static File selectSourceFile() {
    JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
    FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX Files", "xlsx", "XLSX");
    jfc.setAcceptAllFileFilterUsed(false);
    jfc.addChoosableFileFilter(filter);

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
      new ParseResultXLSX(this.sourceFile).process();
    } catch (IOException e) {
      throw new RuntimeException("Exception processing " + this.sourceFile, e);
    }
    // end run
  }

}
