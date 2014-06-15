package org.ccci.gto.pdi.trans.steps.googlespreadsheet;

import org.pentaho.di.core.CheckResult;
import org.pentaho.di.core.CheckResultInterface;
import org.pentaho.di.core.annotations.Step;
import org.pentaho.di.core.database.DatabaseMeta;
import org.pentaho.di.core.exception.KettleException;
import org.pentaho.di.core.exception.KettleStepException;
import org.pentaho.di.core.exception.KettleValueException;
import org.pentaho.di.core.exception.KettleXMLException;
import org.pentaho.di.core.row.RowMetaInterface;
import org.pentaho.di.core.variables.VariableSpace;
import org.pentaho.di.core.xml.XMLHandler;
import org.pentaho.di.repository.ObjectId;
import org.pentaho.di.repository.Repository;
import org.pentaho.di.trans.Trans;
import org.pentaho.di.trans.TransMeta;
import org.pentaho.di.trans.step.*;
import org.pentaho.metastore.api.IMetaStore;
import org.w3c.dom.Node;

import java.security.KeyStore;
import java.util.List;

@Step(id = "GoogleSpreadsheetOutput", image = "google-spreadsheet-output.png", name = "Google Spreadsheet Output",
        description = "Writes to a Google Spreadsheet", categoryDescription = "Output")
public class GoogleSpreadsheetOutputMeta extends BaseStepMeta implements StepMetaInterface {
    private static Class<?> PKG = GoogleSpreadsheetOutputMeta.class;

    private String serviceEmail;
    private KeyStore privateKeyStore;
    private String spreadsheetKey;
    private String worksheetId;

    public GoogleSpreadsheetOutputMeta() {
        super();
    }

    @Override
    public void setDefault() {
        this.serviceEmail = "";
        this.spreadsheetKey = "";
        this.worksheetId = "od6";
        this.privateKeyStore = null;
    }

    @Override
    public String getDialogClassName() {
        return "org.ccci.gto.pdi.ui.trans.steps.googlespreadsheet.GoogleSpreadsheetOutputDialog";
    }

    public String getServiceEmail() {
        return this.serviceEmail == null ? "" : this.serviceEmail;
    }

    public void setServiceEmail(String serviceEmail) {
        this.serviceEmail = serviceEmail;
    }

    public KeyStore getPrivateKeyStore() {
        return this.privateKeyStore;
    }

    public void setPrivateKeyStore(KeyStore pks) {
        this.privateKeyStore = pks;
    }

    public String getSpreadsheetKey() {
        return this.spreadsheetKey == null ? "" : this.spreadsheetKey;
    }

    public void setSpreadsheetKey(String key) {
        this.spreadsheetKey = key;
    }

    public String getWorksheetId() {
        return this.worksheetId == null ? "" : this.worksheetId;
    }

    public void setWorksheetId(String id) {
        this.worksheetId = id;
    }

    @Override
    public Object clone() {
        GoogleSpreadsheetOutputMeta retval = (GoogleSpreadsheetOutputMeta) super.clone();
        retval.setServiceEmail(this.serviceEmail);
        retval.setPrivateKeyStore(this.privateKeyStore);
        retval.setSpreadsheetKey(this.spreadsheetKey);
        retval.setWorksheetId(this.worksheetId);
        return retval;
    }

    @Override
    public String getXML() throws KettleException {
        StringBuilder xml = new StringBuilder();
        try {
            xml.append(XMLHandler.addTagValue("serviceEmail", this.serviceEmail));
            xml.append(XMLHandler.addTagValue("spreadsheetKey", this.spreadsheetKey));
            xml.append(XMLHandler.addTagValue("worksheetId", this.worksheetId));
            xml.append(XMLHandler.openTag("privateKeyStore"));
            xml.append(XMLHandler.buildCDATA(GoogleSpreadsheet.base64EncodePrivateKeyStore(this.privateKeyStore)));
            xml.append(XMLHandler.closeTag("privateKeyStore"));
        } catch (Exception e) {
            throw new KettleValueException("Unable to write step to XML", e);
        }
        return xml.toString();
    }

    @Override
    public void loadXML(Node stepnode, List<DatabaseMeta> databases, IMetaStore metaStore) throws KettleXMLException {
        try {
            this.serviceEmail = XMLHandler.getTagValue(stepnode, "serviceEmail");
            this.spreadsheetKey = XMLHandler.getTagValue(stepnode, "spreadsheetKey");
            this.worksheetId = XMLHandler.getTagValue(stepnode, "worksheetId");
            this.privateKeyStore = GoogleSpreadsheet.base64DecodePrivateKeyStore(XMLHandler.getTagValue(stepnode, "privateKeyStore"));
        } catch (Exception e) {
            throw new KettleXMLException("Unable to load step from XML", e);
        }
    }

    @Override
    public void readRep(Repository rep, IMetaStore metaStore, ObjectId id_step, List<DatabaseMeta> databases) throws KettleException {
        try {
            this.serviceEmail = rep.getStepAttributeString(id_step, "serviceEmail");
            this.spreadsheetKey = rep.getStepAttributeString(id_step, "spreadsheetKey");
            this.worksheetId = rep.getStepAttributeString(id_step, "worksheetId");
            this.privateKeyStore = GoogleSpreadsheet.base64DecodePrivateKeyStore(rep.getStepAttributeString(id_step, "privateKeyStore"));
        } catch (Exception e) {
            throw new KettleException("Unexpected error reading step information from the repository", e);
        }
    }

    @Override
    public void saveRep(Repository rep, IMetaStore metaStore, ObjectId id_transformation, ObjectId id_step) throws KettleException {
        try {
            rep.saveStepAttribute(id_transformation, id_step, "serviceEmail", this.serviceEmail);
            rep.saveStepAttribute(id_transformation, id_step, "spreadsheetKey", this.spreadsheetKey);
            rep.saveStepAttribute(id_transformation, id_step, "worksheetId", this.worksheetId);
            rep.saveStepAttribute(id_transformation, id_step, "privateKeyStore", GoogleSpreadsheet.base64EncodePrivateKeyStore(this.privateKeyStore));
        } catch (Exception e) {
            throw new KettleException("Unable to save step information to the repository for id_step=" + id_step, e);
        }
    }

    @Override
    public void check(List<CheckResultInterface> remarks, TransMeta transMeta, StepMeta stepMeta, RowMetaInterface prev, String[] input, String[] output, RowMetaInterface info, VariableSpace space, Repository repository, IMetaStore metaStore) {
        if (prev == null || prev.size() == 0) {
            remarks.add(new CheckResult(CheckResultInterface.TYPE_RESULT_WARNING, "Not receiving any fields from previous steps!", stepMeta));
        } else {
            remarks.add(new CheckResult(CheckResultInterface.TYPE_RESULT_OK, String.format("Step is connected to previous one, receiving %1$d fields", prev.size()), stepMeta));
        }

        if (input.length > 0) {
            remarks.add( new CheckResult(CheckResultInterface.TYPE_RESULT_OK, "Step is receiving info from other steps", stepMeta) );
        } else {
            remarks.add(new CheckResult(CheckResultInterface.TYPE_RESULT_ERROR, "No input received from other steps!", stepMeta));
        }
    }

    @Override
    public StepInterface getStep(StepMeta stepMeta, StepDataInterface stepDataInterface, int copyNr, TransMeta transMeta, Trans trans) {
        return new GoogleSpreadsheetOutput(stepMeta, stepDataInterface, copyNr, transMeta, trans);
    }

    @Override
    public StepDataInterface getStepData() {
        return new GoogleSpreadsheetOutputData();
    }
}
