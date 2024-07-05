package com.knistr.filialenconverter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Converter {

    // INPUT XLSX
    // [0] Übergeordnete Gruppe, [1] Name, [2] Externe ID, [3] Straße, [4] PLZ, [5] Ort, [6] TelNr, [7] eMail

    // OUTPUT CSV
    // [0] Filialnummer, [1] Filiale, [2] Strasse, [3] houseNumber, [4] Country, [5] Postleitzahl, [6] Ort, [7] Vorname, [8] Nachname, [9] Telefonnummer, [10] Email-Kasse
    private final int OUTPUT_CELLS = 11;

    private final List<String[]> input;
    private final List<String[]> output;

    public Converter(final String INPUT_PATH) {
        final String OUTPUT_PATH = "CSV-" + INPUT_PATH.split("\\.")[0] + ".csv";

        input = new ArrayList<>();
        loadInputData(INPUT_PATH);

        output = new ArrayList<>();
        output.add(new String[] { "Filialnummer", "Filiale", "Strasse", "houseNumber", "Country", "Postleitzahl", "Ort", "Vorname", "Nachname", "Telefonnummer", "Email-Kasse" });
        buildOutputData();

        writeOutputFile(OUTPUT_PATH);
    }

    private void loadInputData(final String INPUT_PATH) {
        try (FileInputStream inputStream = new FileInputStream(new File(INPUT_PATH));
             XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {

            XSSFSheet sheet = workbook.getSheetAt(0);

            XSSFTable table = sheet.getTables().get(0);
            CellReference start = table.getStartCellReference();
            CellReference end = table.getEndCellReference();

            // Iterate through Rows & Cells
            for (int r = start.getRow() + 1; r <= end.getRow(); r++) {
                boolean empty = true;
                String[] line = new String[end.getCol() + 1];

                for (int c = start.getCol(); c <= end.getCol(); c++) {
                    String value = getCellValue(sheet.getRow(r).getCell(c));
                    line[c] = value;
                    if (!value.isEmpty()) empty = false;
                }

                // Rows with atleast 1 Value are added
                if (!empty) input.add(line);
            }
        } catch (IOException e) {
            System.err.println(e.getMessage());
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue().toString();
                else return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    private void buildOutputData() {
        for (String[] line : input) {
            String[] outLine = new String[OUTPUT_CELLS];

            // Filialnummer & Filiale
            /*String filialnummer = "";
            String filiale = "";

            String name = line[1];
            char[] nameChars = name.toCharArray();
            for (int i = 0; i < nameChars.length; i++) {
                int next = i + 1;
                if (nameChars[i] == ' ' && next < nameChars.length) {
                    if (isNumeric(nameChars[next])) {
                        filialnummer = name.substring(next);
                        filiale = name.substring(0, i);
                    }
                }
            }*/

            outLine[0] = line[2];
            outLine[1] = line[1];

            // Strasse & houseNumber
            String[] streetAndHousenumber = splitStreetAndHousenumber(line[3]);
            outLine[2] = streetAndHousenumber[0];
            outLine[3] = streetAndHousenumber[1];

            // Country
            // Currently Hardcoded because expected .xlsx only indirectly contains this Information!
            outLine[4] = "DE";

            // Postleitzahl
            outLine[5] = line[4];

            // Ort
            outLine[6] = line[5];

            // Vorname
            outLine[7] = "Vorname";

            // Nachname
            outLine[8] = "Nachname";

            // Telefonnummer
            outLine[9] = line[6];

            // Email-Kasse
            outLine[10] = line[7];

            output.add(outLine);
        }
    }

    private void writeOutputFile(final String OUTPUT_PATH) {
        try (FileWriter writer = new FileWriter(OUTPUT_PATH)) {
            StringBuilder csv = new StringBuilder();

            for (String[] line : output) {
                for (int c = 0; c < line.length; c++) {
                    csv.append(line[c]);
                    if (c < (line.length - 1)) csv.append(",");
                    else csv.append("\n");
                }
            }

            writer.write(csv.toString());
            System.out.println("Done!");
        } catch (IOException e) {
            System.err.println(e.getMessage());
        }
    }

    private String[] splitStreetAndHousenumber(String input) {
        char[] chars = input.toCharArray();

        // Iterate through every Char
        for (int i = 0; i < chars.length; i++) {
            if (chars[i] == ' ' && (i + 1) < chars.length) {
                int next = i + 1;

                // Fires on Char-Combination: ' ' + '{numeric}'
                if (isNumeric(chars[next])) {
                    int j = next + 1;

                    boolean cont = false;

                    // Check if ecountered Number belongs to the Street-Name
                    // By checking for the same Char-Combination a Second Time
                    while (j < chars.length) {
                        if (chars[j] == ' ' && (j + 1) < chars.length) {
                            if (isNumeric(chars[j + 1])) {
                                // Update i, break while-loop, contiue after
                                i = j - 1;
                                cont = true;
                                break;
                            }
                        }

                        j++;
                    }

                    if (cont) continue;

                    // Return String[2] { Street, houseNumber}
                    return new String[] { input.substring(0, i).replace(",", ""), input.substring(next) };
                }
            }
        }

        return new String[] { input, "" };
    }

    private boolean isNumeric(char c) {
        return c == '0' || c == '1' || c == '2' || c == '3' || c == '4' || c == '5' || c == '6' || c == '7' || c == '8' || c == '9';
    }

}