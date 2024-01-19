package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ooxml.POIXMLException;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

public class LazySupplierTest {
    @Test
    public void testLazySupplier() {
        LazySupplier<String> lazySupplier = new LazySupplier<>(() -> "test");
        assertEquals("test", lazySupplier.getContent());
    }

    @Test
    public void testLazySupplierException() {
        LazySupplier<String> lazySupplier = new LazySupplier<>(() -> {
            throw new POIXMLException("test-exception");
        });
        try {
            assertEquals("test", lazySupplier.getContent());
            fail("expected exception");
        } catch (POIXMLException e) {
            assertEquals("test-exception", e.getMessage());
        }
    }
}
