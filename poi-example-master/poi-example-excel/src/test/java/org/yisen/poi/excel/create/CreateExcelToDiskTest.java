package org.yisen.poi.excel.create;


import org.junit.Test;
import org.yisen.poi.excel.create.CreateExcelToDisk;

public class CreateExcelToDiskTest {
	@Test
	public void testCreateSimpleExcelToDisk() throws Exception {
		CreateExcelToDisk.createSimpleExcelToDisk();
	}
	
	@Test
	public void testCreateMergedRegionExcelToDisk() throws Exception {
		CreateExcelToDisk.createMergedRegionExcelToDisk();;
	}
	
}
