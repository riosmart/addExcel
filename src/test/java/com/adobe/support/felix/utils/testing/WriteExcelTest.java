package com.adobe.support.felix.utils.testing;

import static org.junit.Assert.*;

import java.io.IOException;


import org.junit.Test;

import com.adobe.support.felix.utils.WriteExcel;

public class WriteExcelTest {
	
	@Test
	public void test() throws IOException {
		double output = WriteExcel.calcCells(1); //30
		assertEquals(30.0, output,0);
	}

}
