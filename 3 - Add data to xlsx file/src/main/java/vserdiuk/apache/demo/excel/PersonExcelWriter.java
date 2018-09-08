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

    /**
     * The method prepares a Workbook by preparing the header row and the content area
     *
     * @param personList
     * @return Workbook
     */
    private Workbook prepareWorkbook(List<Person> personList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet personSheet = workbook.createSheet("Persons");
        prepareHeader(personSheet);
        prapareTable(personSheet, personList);
        return workbook;
    }

    /**
     * The method prepares person sheet
     *
     * @param personSheet
     */
    private void prepareHeader(Sheet personSheet) {
        //create header row
        Row headerRow = personSheet.createRow(0);

        //create cells
        Cell firstNameHeaderCell = headerRow.createCell(0);
        Cell lastNameHeaderCell = headerRow.createCell(1);
        Cell birthdayHeaderCell = headerRow.createCell(2);
        Cell emailHeaderCell = headerRow.createCell(3);
        Cell phoneNumberHeaderCell = headerRow.createCell(4);
        Cell marriedHeaderCell = headerRow.createCell(5);

        //set cells values
        firstNameHeaderCell.setCellValue("First name");
        lastNameHeaderCell.setCellValue("Last name");
        birthdayHeaderCell.setCellValue("Birthday");
        emailHeaderCell.setCellValue("Email");
        phoneNumberHeaderCell.setCellValue("Phone number");
        marriedHeaderCell.setCellValue("Married");
    }

    /**
     * The method prepares person content area - lines with Person's data
     * (First name, Last name, Birthday, Email, Phone number, Married)
     *
     * @param personSheet
     * @param personList
     */
    private void prapareTable(Sheet personSheet, List<Person> personList) {
        for (int i=0; i<personList.size(); i++) {
            //create header row
            Row row = personSheet.createRow(i + 1);

            //create cells
            Cell firstNameCell = row.createCell(0);
            Cell lastNameCell = row.createCell(1);
            Cell birthdayCell = row.createCell(2);
            Cell emailCell = row.createCell(3);
            Cell phoneNumberCell = row.createCell(4);
            Cell marriedCell = row.createCell(5);

            //set cells values
            firstNameCell.setCellValue(personList.get(i).getFirstName());
            lastNameCell.setCellValue(personList.get(i).getLastName());
            birthdayCell.setCellValue(personList.get(i).getBirthday().toString());
            emailCell.setCellValue(personList.get(i).getEmail());
            phoneNumberCell.setCellValue(personList.get(i).getPhoneNumber());
            marriedCell.setCellValue(personList.get(i).isMarried());
        }
    }

}














