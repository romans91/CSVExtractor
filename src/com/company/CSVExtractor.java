package com.company;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Iterator;
import java.util.NoSuchElementException;

public class CSVExtractor {

    public static void main(String[] args) throws IOException {
        // write your code here
        if (args.length == 0) {
            args = new String[] { "Data.xlsx" };
        }

        for (String inXlsxFilename : args){
            if (Files.exists(Paths.get(inXlsxFilename))) {
                System.out.println(String.format("%s found.\n", inXlsxFilename));
                extractCsvs(inXlsxFilename);
            } else {
                System.out.println(String.format("%s does not exist in this directory.\n", inXlsxFilename));
            }
        }
    }

    public static void extractCsvs(String inXlsxFilename) throws IOException {
        boolean isExtractingCsv = false;
        String outCsvFilename = "";
        int outCsvWidth = 0;
        StringBuilder outCsvContent = new StringBuilder();

        PrintWriter pw = new PrintWriter("PW");
        pw.close();
        Files.delete(Paths.get("PW"));

        Iterator<Sheet> sheets = new XSSFWorkbook(new FileInputStream(new File(inXlsxFilename))).sheetIterator();

        while (sheets.hasNext()) {
            Sheet sheet = sheets.next();
            Iterator<Row> rows = sheet.rowIterator();

            while (rows.hasNext()) {
                Row row = rows.next();
                Iterator<Cell> cells = row.cellIterator();

                while (cells.hasNext()) {
                    Cell cell = cells.next();

                    if (isExtractingCsv) {
                        if (cell.getColumnIndex() <= outCsvWidth) {
                            if (cell.getColumnIndex() == 0 && cell.toString().equals(".")) {
                                isExtractingCsv = false;
                                System.out.println("\tLength: " + outCsvContent.toString().split("\r\n").length + '\n');

                                pw.write(outCsvContent.toString());
                                pw.close();
                            } else {
                                switch (cell.getCellTypeEnum()) {
                                    case NUMERIC:
                                        outCsvContent.append(cell.getNumericCellValue() % 1 == 0 ? (int)cell.getNumericCellValue() : cell);
                                        break;
                                    case FORMULA:
                                        outCsvContent.append(cell.getNumericCellValue());
                                        break;
                                    default:
                                        outCsvContent.append(cell);
                                        break;
                                }
                            }

                            outCsvContent.append(cell.getColumnIndex() == outCsvWidth ? "\r\n" : ",");
                        }
                    } else {
                        if (cell.getColumnIndex() == 0 && cell.toString().endsWith(".csv")) {
                            isExtractingCsv = true;
                            String[] dirs = inXlsxFilename.split("/");
                            String dirPath = Arrays.toString(Arrays.copyOf(dirs, dirs.length - 1));
                            dirPath = dirPath.replace("]", "")
                                    .replace("[", "")
                                    .replace(", ", "/");

                            if (dirPath.length() > 0) {
                                dirPath += "/";
                            }

                            outCsvFilename = dirPath + "csv/" + cell.toString();

                            if (!Files.isDirectory(Paths.get(dirPath + "csv/"))) {
                                Files.createDirectory(Paths.get(dirPath + "csv/"));
                            }

                            pw = new PrintWriter(outCsvFilename);

                            try {
                                outCsvWidth = (int) cells.next().getNumericCellValue() - 1;
                            } catch (NoSuchElementException e) {
                                outCsvWidth = -1;
                                System.out.println("\tNo CSV width specified");
                            } catch (IllegalStateException e ) {
                                outCsvWidth = -1;
                                System.out.println("\tCSV width is not numeric");
                            }

                            System.out.println("Filename: " + outCsvFilename);
                            System.out.println(outCsvWidth < 0 ? "\tCSV width specifier must be numeric & adjacent to the filename in column A" : "\tWidth: " + (outCsvWidth + 1));
                            outCsvContent.setLength(0);
                        }
                    }
                }
            }
        }

        if (isExtractingCsv) {
            System.out.println("\tReached the end of the spreadsheet before the end of the CSV\nCSV not created");
            pw.close();
            Files.delete(Paths.get(outCsvFilename));
        }
    }
}
