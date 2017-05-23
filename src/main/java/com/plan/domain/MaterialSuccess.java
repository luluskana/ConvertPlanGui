package com.plan.domain;

import java.io.Serializable;

public class MaterialSuccess implements Serializable {

    private String foamtecPart;

    private String customerPart;

    private int january;

    private int february;

    private int march;

    private int april;

    private int may;

    private int june;

    private int july;

    private int august;

    private int september;

    private int october;

    private int november;

    private int december;

    private int future;

    private int total;

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

    public int getJanuary() {
        return january;
    }

    public void setJanuary(int january) {
        this.january = january;
    }

    public int getFebruary() {
        return february;
    }

    public void setFebruary(int february) {
        this.february = february;
    }

    public int getMarch() {
        return march;
    }

    public void setMarch(int march) {
        this.march = march;
    }

    public int getApril() {
        return april;
    }

    public void setApril(int april) {
        this.april = april;
    }

    public int getMay() {
        return may;
    }

    public void setMay(int may) {
        this.may = may;
    }

    public int getJune() {
        return june;
    }

    public void setJune(int june) {
        this.june = june;
    }

    public int getJuly() {
        return july;
    }

    public void setJuly(int july) {
        this.july = july;
    }

    public int getAugust() {
        return august;
    }

    public void setAugust(int august) {
        this.august = august;
    }

    public int getSeptember() {
        return september;
    }

    public void setSeptember(int september) {
        this.september = september;
    }

    public int getOctober() {
        return october;
    }

    public void setOctober(int october) {
        this.october = october;
    }

    public int getNovember() {
        return november;
    }

    public void setNovember(int november) {
        this.november = november;
    }

    public int getDecember() {
        return december;
    }

    public void setDecember(int december) {
        this.december = december;
    }

    public int getFuture() {
        return future;
    }

    public void setFuture(int future) {
        this.future = future;
    }

    public int getTotal() {
        return total;
    }

    public void setTotal(int total) {
        this.total = total;
    }

    @Override
    public String toString() {
        return "MaterialSuccess{" +
                "foamtecPart='" + foamtecPart + '\'' +
                ", customerPart='" + customerPart + '\'' +
                ", january=" + january +
                ", february=" + february +
                ", march=" + march +
                ", april=" + april +
                ", may=" + may +
                ", june=" + june +
                ", july=" + july +
                ", august=" + august +
                ", september=" + september +
                ", october=" + october +
                ", november=" + november +
                ", december=" + december +
                ", future=" + future +
                ", total=" + total +
                '}';
    }
}
