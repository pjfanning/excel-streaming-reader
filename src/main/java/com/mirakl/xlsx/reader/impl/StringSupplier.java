package com.mirakl.xlsx.reader.impl;

class StringSupplier implements Supplier {
    private final String val;

    StringSupplier(String val) {
        this.val = val;
    }

    @Override
    public Object getContent() {
        return val;
    }
}
