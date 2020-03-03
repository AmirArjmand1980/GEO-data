package com.amirarjmand.pgd.Inspection_Frags;

public class SenRecData_Model {

    public SenRecData_Model(String qtyUseSensor, String qtySpareSensor, String senType) {
        this.qtyUseSensor = qtyUseSensor;
        this.qtySpareSensor = qtySpareSensor;
        SenType = senType;
    }

    public String getQtyUseSensor() {
        return qtyUseSensor;
    }

    public void setQtyUseSensor(String qtyUseSensor) {
        this.qtyUseSensor = qtyUseSensor;
    }

    public String getQtySpareSensor() {
        return qtySpareSensor;
    }

    public void setQtySpareSensor(String qtySpareSensor) {
        this.qtySpareSensor = qtySpareSensor;
    }

    public String getSenType() {
        return SenType;
    }

    public void setSenType(String senType) {
        SenType = senType;
    }

    public String qtyUseSensor;
    public String qtySpareSensor;
    public String SenType;


}
