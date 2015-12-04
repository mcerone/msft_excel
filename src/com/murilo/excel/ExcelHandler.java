/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.murilo.excel;

import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author munascim
 */
public class ExcelHandler {
    
    private Workbook wb = null;
    private int numSheets=0;
    private String fileName=null;
    private int[] numRowsPerSheet = null;
    private String[] sheetsNames = null;
    
    public ExcelHandler(String fileName){
        try {
            this.fileName=fileName;
            wb = newWorkbook(fileName);
            setAttributes();
        } catch (IOException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ExcelHandler.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void setAttributes(){
        numSheets = wb.getNumberOfSheets();
        numRowsPerSheet = new int[numSheets];
        sheetsNames = new String[numSheets];
        for (int i = 0; i < numSheets; i++) {
            Sheet sht = wb.getSheetAt(i);
            sheetsNames[i] = sht.getSheetName();
            numRowsPerSheet[i] = sht.getPhysicalNumberOfRows();
        }
    }
    
    public void printExcelFileSummary(){
        System.out.println("Summary of "+fileName+": ");
        System.out.println("\t"+numSheets+" sheets");
        for (int i = 0; i < numSheets; i++) {
            System.out.println("\t\t"+sheetsNames[i]+" has "+numRowsPerSheet[i]+" rows.");            
        }
    }
    
    private Workbook newWorkbook(String filename) throws IOException, InvalidFormatException{
        
        Workbook wb = null;
           
        if(filename.toLowerCase().endsWith("xls")){
            wb = (HSSFWorkbook) WorkbookFactory.create(new File(filename));
        }
        if(filename.toLowerCase().endsWith("xlsx")){
            wb = (XSSFWorkbook) WorkbookFactory.create(new File(filename));
        }
        return wb;
    }
}
