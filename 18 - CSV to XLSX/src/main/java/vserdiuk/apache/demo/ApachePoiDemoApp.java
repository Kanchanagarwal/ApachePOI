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

import java.io.*;

public class ApachePoiDemoApp {
    public static void main(String[] args) {
        try (FileOutputStream outputStream = new FileOutputStream("demo.xlsx")) {
            ApachePoiDemoApp app = new ApachePoiDemoApp();
            File file = app.getFile("MOCK_DATA.csv");

            FileInputStream fileInputStream = new FileInputStream(file);
            BufferedReader reader = new BufferedReader(new InputStreamReader(fileInputStream));

            Workbook workbook = new XSSFWorkbook();
            XSSFSheet personSheet = (XSSFSheet) workbook.createSheet("Persons");

            String csvFileLine = "";
            int rowNomber = 0;
            while ((csvFileLine = reader.readLine()) != null ){
                String[] columns = csvFileLine.split(",");

                Row row = personSheet.createRow(rowNomber);
                for (int i=0; i<columns.length; i++){
                    Cell cell = row.createCell(i);
                    cell.setCellValue(columns[i]);
                }
                rowNomber++;
            }

            workbook.write(outputStream);
            System.out.println("An Excel file " + "test.xlsx" + " has been created");
            workbook.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }

    }

    private File getFile(String fileName) {
        ClassLoader classLoader = getClass().getClassLoader();
        File file = new File(classLoader.getResource(fileName).getFile());
        return file;
    }

}
