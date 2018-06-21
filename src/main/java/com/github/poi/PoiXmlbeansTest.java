/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package com.github.poi;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiXmlbeansTest {

    private static String unicodeText = "𝝊𝝋𝝌𝝍𝝎𝝏𝝐𝝑𝝒𝝓𝝔𝝕𝝖𝝗𝝘𝝙𝝚𝝛𝝜𝝝𝝞𝝟𝝠𝝡𝝢𝝣𝝤𝝥𝝦𝝧𝝨𝝩𝝪𝝫𝝬𝝭𝝮𝝯𝝰𝝱𝝲𝝳𝝴𝝵𝝶𝝷𝝸𝝹𝝺";

    public static void main(String[] args) {
        File tf = null;
        try {
            try(Workbook wb = new XSSFWorkbook()) {
                addCells(wb);
            }
            try(Workbook wb = new SXSSFWorkbook()) {
                addCells(wb);
            }
            String filename = args.length > 0 ? args[0] : "sample.xlsx";
            try (FileInputStream fis = new FileInputStream(filename);
                    XSSFWorkbook wb = new XSSFWorkbook(fis)) {
                System.out.println("Reading " + filename);
                printText(wb);
                try(Workbook wb2 = writeOutAndReadBack(wb)) {
                    System.out.println();
                    System.out.println("Content after writeOutAndReadBack");
                    printText(wb2);
                }
            }
        } catch (Throwable t) {
            t.printStackTrace();
        }
    }

    private static void addCells(Workbook wb) throws IOException {
        File tf = TempFile.createTempFile("poi-xmlbeans-test", ".xlsx");
        try {
            Sheet sheet = wb.createSheet("Sheet1");
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue(unicodeText);
            try (FileOutputStream os = new FileOutputStream(tf)) {
                wb.write(os);
            }
            try (FileInputStream fis = new FileInputStream(tf);
                 XSSFWorkbook wb2 = new XSSFWorkbook(fis)) {
                System.out.println("Testing setCellValue");
                printText(wb2);
            }
        } finally {
            tf.delete();
        }
    }
    
    private static <R extends Workbook> R writeOutAndReadBack(R wb) {
        Workbook result;
        try {
            result = readBack(writeOut(wb));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        @SuppressWarnings("unchecked")
        R r = (R) result;
        return r;
    }
    
    private static XSSFWorkbook readBack(ByteArrayOutputStream out) throws IOException {
        InputStream is = new ByteArrayInputStream(out.toByteArray());
        out.close();
        try {
            return new XSSFWorkbook(is);
        } finally {
            is.close();
        }
    }
    
    private static <R extends Workbook> ByteArrayOutputStream writeOut(R wb) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream(8192);
        wb.write(out);
        return out;
    }
    
    
    private static void printText(Workbook wb) {
        for(Iterator<Sheet> sheetIter = wb.sheetIterator(); sheetIter.hasNext();) {
            Sheet sheet = sheetIter.next();
            System.out.println(sheet.getSheetName());
            for(Iterator<Row> rowIter = sheet.rowIterator(); rowIter.hasNext(); ) {
                Row row = rowIter.next();
                StringBuilder sb = new StringBuilder();
                boolean first = true;
                for(Iterator<Cell> cellIter = row.cellIterator(); cellIter.hasNext(); ) {
                    if(first) {
                        first = false;
                    } else {
                        sb.append(',');
                    }
                    Cell cell = cellIter.next();
                    sb.append(cell.toString());
                }
                System.out.println(sb.toString());
            }
        }
    }

}
