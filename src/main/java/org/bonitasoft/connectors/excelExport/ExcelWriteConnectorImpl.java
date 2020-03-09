package org.bonitasoft.connectors.excelExport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bonitasoft.engine.connector.ConnectorException;

import java.io.FileOutputStream;
import java.util.List;

/**
 * The connector execution will follow the steps
 * 1 - setInputParameters() --> the connector receives input parameters values
 * 2 - validateInputParameters() --> the connector can validate input parameters values
 * 3 - connect() --> the connector can establish a connection to a remote server (if necessary)
 * 4 - executeBusinessLogic() --> execute the connector
 * 5 - getOutputParameters() --> output are retrieved from connector
 * 6 - disconnect() --> the connector can close connection to remote server (if any)
 */
public class ExcelWriteConnectorImpl extends AbstractExcelWriteConnectorImpl {

    @Override
    protected void executeBusinessLogic() throws ConnectorException {
        //Get access to the connector input parameters
        List<String> headers = getHeader();
        List<Object> data = getData();

        //TODO execute your business logic here

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(getSheetName());
        int rowCount = 0;
        Row row = sheet.createRow(rowCount++);
        int columnCount = 0;
        for (String header : headers) {
            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(header);
        }

        for (Object datum : data) {

            row = sheet.createRow(rowCount++);
            columnCount = 0;
            datum
            rowData.each {
                cellData ->
                        Cell cell = row.createCell(columnCount++);
                if (cellData instanceof String) {
                    cell.setCellValue((String) cellData);
                } else if (cellData instanceof Integer) {
                    cell.setCellValue((Integer) cellData);
                }
            }
        }


        data.each {
            rowData ->
                    row = sheet.createRow(rowCount++);
            columnCount = 0;
            rowData.each {
                cellData ->
                        Cell cell = row.createCell(columnCount++);
                if (cellData instanceof String) {
                    cell.setCellValue((String) cellData);
                } else if (cellData instanceof Integer) {
                    cell.setCellValue((Integer) cellData);
                }
            }
        }
        FileOutputStream outputStream = new FileOutputStream(file);
        try {
            outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            workbook.close();
        } finally {
            outputStream.close();
        }
        //WARNING : Set the output of the connector execution. If outputs are not set, connector fails
        //setExcelFile(excelFile);

    }

    @Override
    public void connect() throws ConnectorException {
        //[Optional] Open a connection to remote server

    }

    @Override
    public void disconnect() throws ConnectorException {
        //[Optional] Close connection to remote server

    }

}
