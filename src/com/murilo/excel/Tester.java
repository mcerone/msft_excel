/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.murilo.excel;

/**
 *
 * @author munascim
 */
public class Tester {
    
    public static void main(String[] args) {
        ExcelHandler eh = new ExcelHandler("C:\\temp\\Amx.xlsx");
        eh.printExcelFileSummary();
    }
    
}
