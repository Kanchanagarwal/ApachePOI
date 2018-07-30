package vserdiuk.apache.demo.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import vserdiuk.apache.demo.model.Person;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.ZoneId;
import java.util.Date;
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
        XSSFSheet personSheet = (XSSFSheet) workbook.createSheet("Persons");

        personSheet.setZoom(200); // set sheet zoom to 200 points
        prepareHeader(personSheet);
        prapareTable(personSheet, personList);

        int firstRowIndex = personSheet.getFirstRowNum();
        int lastRowIndex = personSheet.getLastRowNum();
        int firstColumnIndex = personSheet.getRow(0).getFirstCellNum();
        int lastColumnIndex = personSheet.getRow(0).getLastCellNum()-1;

        personSheet.setAutoFilter(new CellRangeAddress(
                firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex));

        setAutoSizeColumn(personSheet);

        return workbook;
    }

    private void setAutoSizeColumn(Sheet personSheet) {
        int columnCount = personSheet.getRow(0).getLastCellNum()-1;
        for (int i=0; i<columnCount; i++) {
            personSheet.autoSizeColumn(i); // set the auto size for a column according to a column content
        }
    }

    private void prepareHeader(Sheet personSheet) {
        //create row
        Row headerRow = personSheet.createRow(0);

        //create cell style
        CellStyle cellStyle = getHeaderCellStyle(personSheet);

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

        //add style to cells
        firstNameHeaderCell.setCellStyle(cellStyle);
        lastNameHeaderCell.setCellStyle(cellStyle);
        birthdayHeaderCell.setCellStyle(cellStyle);
        emailHeaderCell.setCellStyle(cellStyle);
        phoneNumberHeaderCell.setCellStyle(cellStyle);
        marriedHeaderCell.setCellStyle(cellStyle);
    }

    private void prapareTable(Sheet personSheet, List<Person> personList) {
        for (int i=0; i<personList.size(); i++) {
            //create row
            Row row = personSheet.createRow(i + 1);

            //create common cell style
            CellStyle commonStyle = getTableCellStyle(personSheet);

            CellStyle textStyle = getTextStyle(personSheet);
            CellStyle dateStyle = getDateStyle(personSheet);
            CellStyle phoneStyle = getPhoneStyle(personSheet);

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
            birthdayCell.setCellValue(Date.from(personList.get(i).getBirthday()
                    .atStartOfDay(ZoneId.systemDefault()).toInstant()));
            emailCell.setCellValue(personList.get(i).getEmail());
            phoneNumberCell.setCellValue(personList.get(i).getPhoneNumber());
            marriedCell.setCellValue(personList.get(i).isMarried());

            //add style to cells
            firstNameCell.setCellStyle(textStyle);
            lastNameCell.setCellStyle(textStyle);
            birthdayCell.setCellStyle(dateStyle);
            emailCell.setCellStyle(textStyle);
            phoneNumberCell.setCellStyle(phoneStyle);
            marriedCell.setCellStyle(commonStyle);
        }
    }

    private CellStyle getHeaderCellStyle(Sheet personSheet) {
        Font font = personSheet.getWorkbook().createFont();
        font.setBold(true); //setting font style as bold
        font.setFontName("Arial"); //setting font name as Arial
        font.setFontHeightInPoints((short) 14); //setting font size
        font.setColor(IndexedColors.WHITE.getIndex()); //setting font color

        CellStyle cellStyle = personSheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER); //setting horizontal alignment to center
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //setting vertical alignment to center
        cellStyle.setBorderTop(BorderStyle.MEDIUM); //setting border style to medium (border color is black by default)
        cellStyle.setBorderRight(BorderStyle.MEDIUM);
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex()); //setting cell background to blue
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); //setting type of cell background filling
        cellStyle.setFont(font);
        return cellStyle;
    }

    private CellStyle getTableCellStyle(Sheet personSheet) {
        Font font = personSheet.getWorkbook().createFont();
        font.setItalic(true);
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);

        CellStyle cellStyle = personSheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setTopBorderColor(IndexedColors.ORANGE.getIndex()); //setting border color to orange
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.ORANGE.getIndex());
        cellStyle.setRightBorderColor(IndexedColors.ORANGE.getIndex());
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.ORANGE.getIndex());
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.ORANGE.getIndex());
        cellStyle.setBorderLeft(BorderStyle.THIN);

        return cellStyle;
    }

    private CellStyle getTextStyle(Sheet personSheet) {
        CellStyle cellStyle = getTableCellStyle(personSheet);
        DataFormat dataFormat = personSheet.getWorkbook().createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("@"));
        return cellStyle;
    }

    private CellStyle getDateStyle(Sheet personSheet) {
        CellStyle cellStyle = getTableCellStyle(personSheet);

        CreationHelper createHelper = personSheet.getWorkbook().
                getCreationHelper();

        cellStyle.setDataFormat(createHelper.createDataFormat().
                getFormat("MMMM dd, yyyy"));

        return cellStyle;
    }

    private CellStyle getPhoneStyle(Sheet personSheet) {
        CellStyle cellStyle = getTableCellStyle(personSheet);
        CreationHelper createHelper = personSheet.getWorkbook().getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat()
                .getFormat("(###) ###-####"));
        return cellStyle;
    }

}










