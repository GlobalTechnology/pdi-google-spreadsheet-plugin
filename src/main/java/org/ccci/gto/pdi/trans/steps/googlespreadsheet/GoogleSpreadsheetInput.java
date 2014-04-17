package org.ccci.gto.pdi.trans.steps.googlespreadsheet;


import com.google.gdata.client.spreadsheet.FeedURLFactory;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.*;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.row.RowDataUtil;
import org.pentaho.di.core.row.RowMeta;
import org.pentaho.di.core.row.ValueMetaInterface;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.*;

import java.net.URL;

public class GoogleSpreadsheetInput extends BaseStep implements StepInterface {

    private GoogleSpreadsheetInputMeta meta;
    private GoogleSpreadsheetInputData data;

    public GoogleSpreadsheetInput(StepMeta meta, StepDataInterface data, int num, TransMeta transMeta, Trans trans) {
        super(meta, data, num, transMeta, trans);
    }

    @Override
    public boolean init(StepMetaInterface smi, StepDataInterface sdi) {
        meta = (GoogleSpreadsheetInputMeta) smi;
        data = (GoogleSpreadsheetInputData) sdi;

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

                URL listFeedURL = FeedURLFactory.getDefault().getListFeedUrl(meta.getSpreadsheetKey(), meta.getWorksheetId(), "private", "full");
                ListFeed listFeed = data.service.getFeed(listFeedURL, ListFeed.class);
                data.rows = listFeed.getEntries();
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
        meta = (GoogleSpreadsheetInputMeta) smi;
        data = (GoogleSpreadsheetInputData) sdi;

        if (first) {
            first = false;
            data.outputRowMeta = new RowMeta();
            meta.getFields(data.outputRowMeta, getStepname(), null, null, this, repository, metaStore);
        }

        try {
            Object[] outputRowData = readRow();
            if (outputRowData == null) {
                setOutputDone();
                return false;
            } else {
                putRow(data.outputRowMeta, outputRowData);
            }
        } catch (Exception e) {
            throw new KettleException(e.getMessage());
        } finally {
            data.currentRow++;
        }
        return true;
    }

    private Object[] readRow() {
        try {
            Object[] outputRowData = RowDataUtil.allocateRowData(data.outputRowMeta.size());
            int outputIndex = 0;

            if (data.currentRow < data.rows.size()) {
                ListEntry row = data.rows.get(data.currentRow);
                for (String tag : row.getCustomElements().getTags()) {
                    String value = row.getCustomElements().getValue(tag);
                    if( value == null )
                        outputRowData[outputIndex++] = null;
                    else
                        outputRowData[outputIndex++] = value.getBytes("UTF-8");
                }
            } else {
                return null;
            }
            return outputRowData;
        } catch (Exception e) {
            return null;
        }
    }
}
