package vserdiuk.apache.demo.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import vserdiuk.apache.demo.model.Person;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class PersonExcelWriter {

    public void write(String fileName, List<Person> personList) {
        Workbook workbook = prepareWorkbook(personList);
        try (FileOutputStream outputStream = new FileOutputStream(fileName)) {
            workbook.write(outputStream);
            System.out.println("An Excel file " + fileName + " has been created");
            workbook.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    private Workbook prepareWorkbook(List<Person> personList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet personSheet = workbook.createSheet("Persons");
        prepareHeader(personSheet);
        prapareTable(personSheet, personList);
        return workbook;
    }

    private void prepareHeader(Sheet personSheet) {
        Row headerRow = personSheet.createRow(0);

        Cell firstNameHeaderCell = headerRow.createCell(0);
        firstNameHeaderCell.setCellValue("First name");

        Cell lastNameHeaderCell = headerRow.createCell(1);
        lastNameHeaderCell.setCellValue("Last name");

        Cell birthdayHeaderCell = headerRow.createCell(2);
        birthdayHeaderCell.setCellValue("Birthday");

        Cell emailHeaderCell = headerRow.createCell(3);
        emailHeaderCell.setCellValue("Email");

        Cell phoneNumberHeaderCell = headerRow.createCell(4);
        phoneNumberHeaderCell.setCellValue("Phone number");

        Cell marriedHeaderCell = headerRow.createCell(5);
        marriedHeaderCell.setCellValue("Married");
    }

    private void prapareTable(Sheet personSheet, List<Person> personList) {
        for (int i=0; i<personList.size(); i++) {
            Row row = personSheet.createRow(i + 1);

            Cell firstNameCell = row.createCell(0);
            firstNameCell.setCellValue(personList.get(i).getFirstName());

            Cell lastNameCell = row.createCell(1);
            lastNameCell.setCellValue(personList.get(i).getLastName());

            Cell birthdayCell = row.createCell(2);
            birthdayCell.setCellValue(personList.get(i).getBirthday().toString());

            Cell emailCell = row.createCell(3);
            emailCell.setCellValue(personList.get(i).getEmail());

            Cell phoneNumberCell = row.createCell(4);
            phoneNumberCell.setCellValue(personList.get(i).getPhoneNumber());

            Cell marriedCell = row.createCell(5);
            marriedCell.setCellValue(personList.get(i).isMarried());
        }
    }

}
