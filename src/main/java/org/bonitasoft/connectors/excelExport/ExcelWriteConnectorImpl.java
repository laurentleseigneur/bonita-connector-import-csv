package org.bonitasoft.connectors.excelExport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bonitasoft.engine.connector.ConnectorException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

public class ExcelWriteConnectorImpl extends AbstractExcelWriteConnectorImpl {

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        //Get access to the connector input parameters
        List<String> headers = getHeader();
        List<List<Object>> data = getData();

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(getSheetName());
        AtomicInteger rowCount = new AtomicInteger();
        Row row = sheet.createRow(rowCount.getAndIncrement());
        AtomicInteger columnCount = new AtomicInteger();
        for (String header1 : headers) {
            Cell cell1 = row.createCell(columnCount.getAndIncrement());
            cell1.setCellValue(header1);
        }

        for (List<Object> dataLine : data) {
            row = sheet.createRow(rowCount.getAndIncrement());
            columnCount = new AtomicInteger();
            for (Object cellData : dataLine) {
                int countAndIncrement = columnCount.getAndIncrement();
                if (cellData != null) {
                    Cell cell = row.createCell(countAndIncrement);
                    String cellClass = cellData.getClass().getSimpleName();
                    switch (cellClass) {
                        case "String":
                            cell.setCellValue((String) cellData);
                            break;
                        case "Double":
                            cell.setCellValue((Double) cellData);
                            break;
                        case "Integer":
                            cell.setCellValue((Integer) cellData);
                            break;
                        case "Long":
                            cell.setCellValue((Long) cellData);
                            break;
                        case "Date":
                            cell.setCellValue((Date) cellData);
                            break;
                        case "LocalDate":
                            cell.setCellValue((LocalDate) cellData);
                            break;
                        case "LocalDateTime":
                            cell.setCellValue((LocalDateTime) cellData);
                            break;
                        case "OffsetDateTime":
                            OffsetDateTime offsetDateTime = (OffsetDateTime) cellData;
                            cell.setCellValue(offsetDateTime.toLocalDateTime());
                            break;
                        default:
                            throw new ConnectorException("cell with type " + cellClass + " is not yet supported." +
                                    " Consider using a String");
                    }
                }


            }
        }

        try {
            File excelFile = null;
            excelFile = File.createTempFile("bonita", "");
            try (FileOutputStream outputStream = new FileOutputStream(excelFile)) {
                workbook.write(outputStream);
                workbook.close();
                setExcelFile(excelFile);
            }
        } catch (IOException e) {
            throw new ConnectorException(e);
        }
    }


}
