package org.ws.excel.service.ifc;

import java.io.File;
import java.util.List;
import java.util.Map;

import org.ws.excel.model.Product;

public interface ExcelService {

	public void createWorkbook(List recordsList,File file);
	
	public void createWorkbookForPOI(List recordsList,File file);
	
	public List<Map<String, String>> wriToExcel(List<Product> list);
}
