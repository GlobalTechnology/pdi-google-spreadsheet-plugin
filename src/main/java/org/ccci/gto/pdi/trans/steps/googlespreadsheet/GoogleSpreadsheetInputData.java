package org.ccci.gto.pdi.trans.steps.googlespreadsheet;

import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.ListEntry;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.trans.step.BaseStepData;
import org.pentaho.di.trans.step.StepDataInterface;

import java.util.List;

class GoogleSpreadsheetInputData extends BaseStepData implements StepDataInterface {

    public String accessToken;
    public SpreadsheetService service;

    public RowMetaInterface outputRowMeta;

    public List<ListEntry> rows;
    public int currentRow = 0;
}
