package io.github.dornol.excelkit.example.app.service;


import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.excel.ExcelHandler;

import java.io.InputStream;

public interface BookService {

    ExcelHandler getExcelHandler();

    CsvHandler getCsvHandler();

    void readExcel(InputStream inputStream);

    void readAndSaveExcel(InputStream inputStream);

    void readCsv(InputStream inputStream);

}
