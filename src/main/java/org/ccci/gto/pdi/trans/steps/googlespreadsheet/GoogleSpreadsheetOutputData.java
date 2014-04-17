package org.ccci.gto.pdi.trans.steps.googlespreadsheet;

import com.google.gdata.client.spreadsheet.FeedURLFactory;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import org.ccci.gto.pdi.trans.steps.googlespreadsheet.spreadsheet.Spreadsheet;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.trans.step.BaseStepData;
import org.pentaho.di.trans.step.StepDataInterface;

import java.net.URL;

class GoogleSpreadsheetOutputData extends BaseStepData implements StepDataInterface {

    public String accessToken;
    public SpreadsheetService service;

    public RowMetaInterface outputRowMeta;

    public Spreadsheet spreadsheet;

    public final FeedURLFactory urlFactory = FeedURLFactory.getDefault();

    public String spreadsheetKey;
    public String worksheetId;

    public URL worksheetFeedURL;
    public URL cellFeedURL;
    public URL cellBatchURL;

    public WorksheetFeed worksheetFeed;
    public CellFeed cellFeed;
}
