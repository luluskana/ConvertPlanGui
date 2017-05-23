package com.plan.domain;

import java.io.Serializable;


public class MaterialConvert implements Serializable {

    private String foamtecPart;

    private String customerPart;

    private String month;

    private int qty;

    public String getFoamtecPart() {
        return foamtecPart;
    }

    public void setFoamtecPart(String foamtecPart) {
        this.foamtecPart = foamtecPart;
    }

    public String getCustomerPart() {
        return customerPart;
    }

    public void setCustomerPart(String customerPart) {
        this.customerPart = customerPart;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public int getQty() {
        return qty;
    }

    public void setQty(int qty) {
        this.qty = qty;
    }
}
