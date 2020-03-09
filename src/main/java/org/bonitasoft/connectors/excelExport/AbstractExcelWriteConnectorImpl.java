package org.bonitasoft.connectors.excelExport;

import org.bonitasoft.engine.connector.AbstractConnector;
import org.bonitasoft.engine.connector.ConnectorValidationException;

import java.util.List;

public abstract class AbstractExcelWriteConnectorImpl extends AbstractConnector {

    protected final static String HEADER_INPUT_PARAMETER = "header";
    protected final static String DATA_INPUT_PARAMETER = "data";
    protected final static String SHEET_NAME_INPUT_PARAMETER = "sheetName";
    protected final String EXCELFILE_OUTPUT_PARAMETER = "excelFile";

    protected final List getHeader() {
        return (List) getInputParameter(HEADER_INPUT_PARAMETER);
    }

    protected final String getSheetName() {
        return (String) getInputParameter(SHEET_NAME_INPUT_PARAMETER);
    }

    protected final List getData() {
        return (List) getInputParameter(DATA_INPUT_PARAMETER);
    }

    protected final void setExcelFile(java.io.File excelFile) {
        setOutputParameter(EXCELFILE_OUTPUT_PARAMETER, excelFile);
    }

    @Override
    public void validateInputParameters() throws ConnectorValidationException {
        try {
            getHeader();
        } catch (ClassCastException cce) {
            throw new ConnectorValidationException("header type is invalid");
        }
        try {
            getData();
        } catch (ClassCastException cce) {
            throw new ConnectorValidationException("data type is invalid");
        }

    }

}
