/*
    Licensed to the Apache Software Foundation (ASF) under one
    or more contributor license agreements.  See the NOTICE file
    distributed with this work for additional information
    regarding copyright ownership.  The ASF licenses this file
    to you under the Apache License, Version 2.0 (the
    "License"); you may not use this file except in compliance
    with the License.  You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing,
    software distributed under the License is distributed on an
    "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
    KIND, either express or implied.  See the License for the
    specific language governing permissions and limitations
    under the License.
*/

package vserdiuk.apache.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ApachePoiDemoApp {
    public static void main(String[] args) {
        try (FileOutputStream outputStream = new FileOutputStream("demo.xlsx")) {
            Workbook workbook = new XSSFWorkbook();
            XSSFSheet personSheet = (XSSFSheet) workbook.createSheet("Demo");

            Row row = personSheet.createRow(0);
            Cell cell1 = row.createCell(0);
            Cell cell2 = row.createCell(1);
            Cell cell3 = row.createCell(2);
            Cell cell4 = row.createCell(3);

            int a = 2;
            int b = 3;
            cell1.setCellValue(a);
            cell2.setCellValue(b);

            String formula = "SUM(" + a + ", " + b + ")";

            cell3.setCellFormula(formula);

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell3);

            cell4.setCellValue(cellValue.getNumberValue());

            workbook.write(outputStream);
            System.out.println("An Excel file " + "test.xlsx" + " has been created");
            workbook.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }

    }
}
