package com.cmg.hipspot.data.jdo;

import com.cmg.hipspot.data.Mirrorable;

import javax.jdo.annotations.PersistenceCapable;
import javax.jdo.annotations.PrimaryKey;

/**
 * Created by lantb on 2014-04-18.
 */
@PersistenceCapable
public class FeedbackModelJDO implements Mirrorable {
    @PrimaryKey
    private String id;

    private String email;

    private String description;

    private String pictureError;

    private String version;

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getPictureError() {
        return pictureError;
    }

    public void setPictureError(String pictureError) {
        this.pictureError = pictureError;
    }

    public String getVersion() {
        return version;
    }

    public void setVersion(String version) {
        this.version = version;
    }

    public String getOsInformation() {
        return osInformation;
    }

    public void setOsInformation(String osInformation) {
        this.osInformation = osInformation;
    }

    public String getStepError() {
        return stepError;
    }

    public void setStepError(String stepError) {
        this.stepError = stepError;
    }

    public String getTestData() {
        return testData;
    }

    public void setTestData(String testData) {
        this.testData = testData;
    }

    private String osInformation;

    private String stepError;

    private String testData;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }
}
