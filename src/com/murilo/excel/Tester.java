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
        ExcelHandler eh = new ExcelHandler("C:\\temp\\HKFS.xlsx");
        eh.printExcelFileSummary();
        int limit = eh.getCountRowsFromSheet(0);
        String linha = "";
        int realLineCounter = 0;
        for (int i = 0; i < limit; i++) {
            linha = eh.getLine(0, i);
            if (linha.trim().equals("")) {
                continue;
            }
            realLineCounter++;
            System.out.println(realLineCounter+": "+linha);
        }
        
    }
    
}
