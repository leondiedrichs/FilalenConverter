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

    /*
    Input .xlsx
    [0] Übergeordnete Gruppe
    [1] Name
    [2] Externe ID
    [3] Straße
    [4] PLZ
    [5] Ort
    [6] TelNr
    [7] eMail
    */

    /*
    Output .csv
    [0] Filialnummer
    [1] Filiale
    [2] Strasse
    [3] houseNumber
    [4] Country
    [5] Postleitzahl
    [6] Ort
    [7] Vorname
    [8] Nachname
    [9] Telefonnummer
    [10] Email-Kasse
    */

    public final int OUTPUT_CELLS = 11;

    private final List<String[]> input;
    private final List<String[]> output;

    public Converter(final String INPUT_PATH) {
        input = new ArrayList<>();
        output = new ArrayList<>();

        loadInput(INPUT_PATH);

        output.add(new String[] { "Filialnummer", "Filiale", "Strasse", "houseNumber", "Country", "Postleitzahl", "Ort", "Vorname", "Nachname", "Telefonnummer", "Email-Kasse" });
        buildOutput();

        writeOutput("CSV-" + INPUT_PATH.split("\\.")[0] + ".csv");
    }

    private void loadInput(final String INPUT_PATH) {
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

    private void buildOutput() {
        for (String[] in : input) {
            String[] out = new String[OUTPUT_CELLS];


            // Filialnummer
            out[0] = in[2];


            // Filiale
            out[1] = in[1];


            String[] streetAndHousenumber = splitStreetAndHousenumber(in[3]);
            // Strasse
            out[2] = streetAndHousenumber[0];
            // houseNumber
            out[3] = streetAndHousenumber[1];


            // Country
            out[4] = "DE";


            // Postleitzahl
            out[5] = in[4];


            // Ort
            out[6] = in[5];


            // Vorname
            out[7] = "Vorname";


            // Nachname
            out[8] = "Nachname";


            // Telefonnummer
            out[9] = in[6].replace(" - ", " ");


            // Email-Kasse
            out[10] = in[7];


            output.add(out);
        }
    }

    private void writeOutput(final String OUTPUT_PATH) {
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
                        if ((chars[j] == ' ' && chars[j - 1] != '-') && (j + 1) < chars.length) {
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