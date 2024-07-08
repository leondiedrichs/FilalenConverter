package com.knistr.filialenconverter;

public class Main {

    private static final String XLSX_FILE_INPUT_PATH;

    // Place File in Root-Directory of Project

    static {
        XLSX_FILE_INPUT_PATH = "Filialverzeichnis.xlsx";
    }

    public static void main(String[] args) {
        new Converter(XLSX_FILE_INPUT_PATH);
    }

}
