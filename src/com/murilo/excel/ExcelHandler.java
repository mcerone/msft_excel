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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
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
    private Sheet[] sheets = null;
    
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
        sheets = new Sheet[numSheets];
        for (int i = 0; i < numSheets; i++) {
            sheets[i] = wb.getSheetAt(i);
            sheetsNames[i] = sheets[i].getSheetName();
            numRowsPerSheet[i] = sheets[i].getPhysicalNumberOfRows();
        }
    }
    
    /*TO DO:
    Methods:
    getHeader from Sheet
    getCell from Sheet
    getCellMetadata from Sheet
    */
    
    public int getNumSheets(){
        return numSheets;
    }
    
    public String getSheetName(int index){
        try{
            return sheetsNames[index];
        }catch (ArrayIndexOutOfBoundsException e){
            System.out.println("There's no sheet with index: "+index);
            return null;
        }
    }
    
    public int getCountRowsFromSheet(int index){
        try{
            return numRowsPerSheet[index];
        }catch (ArrayIndexOutOfBoundsException e){
            System.out.println("There's no sheet with index: "+index);
            return -1;
        }
    }
    
    public void printExcelFileSummary(){
        System.out.println("Summary of "+fileName.substring(fileName.lastIndexOf('\\')+1, fileName.length())+": ");
        System.out.println("\t"+numSheets+" sheets");
        for (int i = 0; i < numSheets; i++) {
            System.out.println("\t\t\""+sheetsNames[i]+"\" has "+numRowsPerSheet[i]+" rows.");            
        }
    }
    
    public String getLine(int sheetIndex,int lineIndex, char separator){
        Row linha = sheets[sheetIndex].getRow(lineIndex);
        String aux = "";
        for (int i = linha.getFirstCellNum(); i < linha.getLastCellNum(); i++) {
            Cell campo = linha.getCell(i);
            if((i+1)!=linha.getLastCellNum()){
                aux = aux + "\""+stringrizeCell(campo)+"\"" + separator; 
            } else{
               aux = aux + "\""+stringrizeCell(campo)+"\"\n";
            }
        }
        
        return aux;
    }
    
    private String stringrizeCell(Cell x){
        
        switch (x.getCellType()){
            case Cell.CELL_TYPE_BLANK:
                return "";
            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(x.getBooleanCellValue());
            case Cell.CELL_TYPE_NUMERIC:
                return String.valueOf(x.getNumericCellValue());
            case Cell.CELL_TYPE_STRING:
                return x.getStringCellValue();
            case Cell.CELL_TYPE_FORMULA:
                switch (x.getCachedFormulaResultType()){
                        case Cell.CELL_TYPE_NUMERIC:
                            return String.valueOf(x.getNumericCellValue());
                        case Cell.CELL_TYPE_STRING:
                            return x.getStringCellValue();
                        case Cell.CELL_TYPE_BLANK:
                            return "";
                }
        }
        
        return null;
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
