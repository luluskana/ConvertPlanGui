package com.plan.domain;

import java.io.Serializable;

public class MaterialPlanning implements Serializable {

    private Long id;

    private String materialFoamtec;

    private String materialCustomer;

    private String materialGroup;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getMaterialFoamtec() {
        return materialFoamtec;
    }

    public void setMaterialFoamtec(String materialFoamtec) {
        this.materialFoamtec = materialFoamtec;
    }

    public String getMaterialCustomer() {
        return materialCustomer;
    }

    public void setMaterialCustomer(String materialCustomer) {
        this.materialCustomer = materialCustomer;
    }

    public String getMaterialGroup() {
        return materialGroup;
    }

    public void setMaterialGroup(String materialGroup) {
        this.materialGroup = materialGroup;
    }
}
