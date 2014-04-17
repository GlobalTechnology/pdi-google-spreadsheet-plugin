package org.ccci.gto.pdi.trans.steps.googlespreadsheet;

import com.google.gdata.client.batch.BatchInterruptedException;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.IFeed;
import com.google.gdata.data.batch.BatchOperationType;
import com.google.gdata.data.batch.BatchStatus;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import com.google.gdata.model.atom.Link;
import com.google.gdata.model.batch.BatchUtils;
import com.google.gdata.util.ServiceException;
import org.ccci.gto.pdi.trans.steps.googlespreadsheet.spreadsheet.Spreadsheet;
import org.ccci.gto.pdi.trans.steps.googlespreadsheet.spreadsheet.SpreadsheetCell;
import org.ccci.gto.pdi.trans.steps.googlespreadsheet.spreadsheet.SpreadsheetRow;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.row.ValueMetaInterface;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.*;

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

class GoogleSpreadsheetOutput extends BaseStep implements StepInterface {

    private GoogleSpreadsheetOutputMeta meta;
    private GoogleSpreadsheetOutputData data;

    public GoogleSpreadsheetOutput(StepMeta meta, StepDataInterface data, int num, TransMeta transMeta, Trans trans) {
        super(meta, data, num, transMeta, trans);
    }

    @Override
    public boolean init(StepMetaInterface smi, StepDataInterface sdi) {
        meta = (GoogleSpreadsheetOutputMeta) smi;
        data = (GoogleSpreadsheetOutputData) sdi;

        if (super.init(smi, sdi)) {
            try {
                data.accessToken = GoogleSpreadsheet.getAccessToken(meta.getServiceEmail(), meta.getPrivateKeyStore());
                if (data.accessToken == null) {
                    logError("Unable to get access token.");
                    setErrors(1L);
                    stopAll();
                    return false;
                }
                data.service = new SpreadsheetService("PentahoKettleTransformStep-v1");
                data.service.setHeader("Authorization", "Bearer " + data.accessToken);

                data.spreadsheetKey = meta.getSpreadsheetKey();
                data.worksheetId = meta.getWorksheetId();

                data.worksheetFeedURL = data.urlFactory.getWorksheetFeedUrl(data.spreadsheetKey, "private", "full");
                data.worksheetFeed = data.service.getFeed(data.worksheetFeedURL, WorksheetFeed.class);

                data.cellFeedURL = data.urlFactory.getCellFeedUrl(data.spreadsheetKey, data.worksheetId, "private",
                        "full");
                data.cellFeed = data.service.getFeed(data.cellFeedURL, CellFeed.class);
                data.cellBatchURL = new URL(data.cellFeed.getLink(Link.Rel.FEED_BATCH, Link.Type.ATOM).getHref());

                data.spreadsheet = new Spreadsheet();
            } catch (Exception e) {
                logError("Error: " + e.getMessage(), e);
                setErrors(1L);
                stopAll();
                return false;
            }

            return true;
        }
        return false;
    }

    @Override
    public synchronized boolean processRow(StepMetaInterface smi, StepDataInterface sdi) throws KettleException {
        meta = (GoogleSpreadsheetOutputMeta) smi;
        data = (GoogleSpreadsheetOutputData) sdi;

        Object[] r = getRow(); // This also waits for a row to be finished.

        if (r != null && first) {
            first = false;
            data.outputRowMeta = getInputRowMeta().clone();
            meta.getFields( data.outputRowMeta, getStepname(), null, null, this, repository, metaStore );

            SpreadsheetRow header = new SpreadsheetRow();
            for (int i = 0; i < data.outputRowMeta.size(); i++) {
                ValueMetaInterface v = data.outputRowMeta.getValueMeta(i);
                header.cells.add(new SpreadsheetCell(v.getName()));
            }
            data.spreadsheet.rows.add(header);
        }

        if (r == null) {

            // Resize the worksheet
            try {
                for (WorksheetEntry worksheet : data.worksheetFeed.getEntries()) {
                    if (worksheet.getCellFeedUrl().equals(data.cellFeedURL)) {
                        worksheet.setColCount(data.spreadsheet.getColumnCount());
                        worksheet.setRowCount(data.spreadsheet.getRowCount());
                        worksheet.update();
                    }
                }

                List<SpreadsheetCell> spreadsheetCells = data.spreadsheet.getCells();
                // Map<String, CellEntry> cellEntries = getCellEntryMap(
                // spreadsheetCells );

                CellFeed batchRequest = new CellFeed();
                for (SpreadsheetCell cell : spreadsheetCells) {
                    // CellEntry batchEntry = new CellEntry( cellEntries.get(
                    // cell.id ) );
                    CellEntry batchEntry = new CellEntry(cell.row, cell.col, cell.value);
                    batchEntry.setId(String.format("%s/%s", data.cellFeedURL.toString(), cell.id));
                    batchEntry.changeInputValueLocal(cell.value);
                    BatchUtils.setBatchId(batchEntry, cell.id);
                    BatchUtils.setBatchOperationType(batchEntry, BatchOperationType.UPDATE);
                    batchRequest.getEntries().add(batchEntry);
                }

                CellFeed batchResponse = batchRequest(data.cellBatchURL, batchRequest);

                for (CellEntry entry : batchResponse.getEntries()) {
                    String batchId = BatchUtils.getBatchId(entry);
                    if (!BatchUtils.isSuccess(entry)) {
                        BatchStatus status = (BatchStatus) BatchUtils.getStatus(entry);
                        logError(String.format("%s failed (%s) %s", batchId, status.getReason(), status.getContent()));
                    }
                }

            } catch (Exception e) {
                throw new KettleException(e);
            }

            setOutputDone();
            return false;
        }

        SpreadsheetRow spreadsheetRow = new SpreadsheetRow();
        for (int i = 0; i < data.outputRowMeta.size(); i++) {
            ValueMetaInterface v = data.outputRowMeta.getValueMeta(i);
            Object value = r[i];
            String str;

            if (v.isNull(value)) {
                str = "";
            } else {
                str = (value instanceof String) ? (String) value : v.getString(value);
            }
            spreadsheetRow.cells.add(new SpreadsheetCell(str));
        }
        data.spreadsheet.rows.add(spreadsheetRow);

        putRow(data.outputRowMeta, r);
        return true;
    }

    private <F extends IFeed> F batchRequest(URL url, F feed) throws IOException,
            ServiceException {
        data.service.setHeader("If-Match", "*");
        F response = data.service.batch(url, feed);
        data.service.setHeader("If-Match", null);
        return response;
    }

    private Map<String, CellEntry> getCellEntryMap(List<SpreadsheetCell> cells) throws
            IOException, ServiceException {

        CellFeed batchRequest = new CellFeed();
        for (SpreadsheetCell cell : cells) {
            CellEntry batchEntry = new CellEntry(cell.row, cell.col, cell.id);
            batchEntry.setId(String.format("%s/%s", data.cellFeedURL.toString(), cell.id));
            BatchUtils.setBatchId(batchEntry, cell.id);
            BatchUtils.setBatchOperationType(batchEntry, BatchOperationType.QUERY);
            batchRequest.getEntries().add(batchEntry);
        }

        CellFeed batchResponse = batchRequest(data.cellBatchURL, batchRequest);

        Map<String, CellEntry> cellEntryMap = new HashMap<String, CellEntry>(cells.size());
        for (CellEntry entry : batchResponse.getEntries()) {
            cellEntryMap.put(BatchUtils.getBatchId(entry), entry);
        }

        return cellEntryMap;
    }
}
