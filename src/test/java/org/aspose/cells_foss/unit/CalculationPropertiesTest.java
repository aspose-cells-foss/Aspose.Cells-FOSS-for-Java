package org.aspose.cells_foss.unit;

import org.aspose.cells_foss.CalculationProperties;
import org.aspose.cells_foss.CellsException;
import org.aspose.cells_foss.Workbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class CalculationPropertiesTest {

    private CalculationProperties props() {
        return new Workbook().getCalculationProperties();
    }

    @Test
    void calculationMode_defaultIsAuto() {
        assertEquals("auto", props().getCalculationMode());
    }

    @Test
    void referenceMode_defaultIsA1() {
        assertEquals("A1", props().getReferenceMode());
    }

    @Test
    void iterate_defaultFalse() {
        assertFalse(props().getIterate());
    }

    @Test
    void iterateCount_defaultOneHundred() {
        assertEquals(100, props().getIterateCount());
    }

    @Test
    void iterateDelta_defaultPointOhOhOne() {
        assertEquals(0.001, props().getIterateDelta(), 1e-10);
    }

    @Test
    void fullPrecision_defaultTrue() {
        assertTrue(props().getFullPrecision());
    }

    @Test
    void calculationCompleted_defaultTrue() {
        assertTrue(props().getCalculationCompleted());
    }

    @Test
    void calculationOnSave_defaultTrue() {
        assertTrue(props().getCalculationOnSave());
    }

    @Test
    void concurrentCalculation_defaultTrue() {
        assertTrue(props().getConcurrentCalculation());
    }

    @Test
    void forceFullCalculation_defaultFalse() {
        assertFalse(props().getForceFullCalculation());
    }

    @Test
    void calculationId_defaultNull() {
        assertNull(props().getCalculationId());
    }

    @Test
    void setCalculationId_negative_throws() {
        CalculationProperties p = props();
        assertThrows(CellsException.class, () -> p.setCalculationId(-1));
    }

    @Test
    void setCalculationId_zero_doesNotThrow() {
        CalculationProperties p = props();
        assertDoesNotThrow(() -> p.setCalculationId(0));
    }

    @Test
    void setIterateCount_negative_throws() {
        CalculationProperties p = props();
        assertThrows(CellsException.class, () -> p.setIterateCount(-1));
    }

    @Test
    void setIterateCount_zero_doesNotThrow() {
        CalculationProperties p = props();
        assertDoesNotThrow(() -> p.setIterateCount(0));
    }

    @Test
    void setIterateDelta_negative_throws() {
        CalculationProperties p = props();
        assertThrows(CellsException.class, () -> p.setIterateDelta(-0.001));
    }

    @Test
    void setIterateDelta_zero_doesNotThrow() {
        CalculationProperties p = props();
        assertDoesNotThrow(() -> p.setIterateDelta(0.0));
    }
}

