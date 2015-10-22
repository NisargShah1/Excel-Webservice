package org.ws.excel.restws.ws;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Response.ResponseBuilder;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.ws.excel.model.Product;
import org.ws.excel.service.ifc.ExcelService;

/**
 * @author NISARG
 *
 */
@Component
@Path("/excel")
public class ExcelExportWs{

	/**
	 * 
	 */

	@Autowired
	private ExcelService excelService;
	
	

	@Path("/jxl/export")
	@GET
	@Produces (MediaType.APPLICATION_OCTET_STREAM)
	@Consumes(MediaType.APPLICATION_JSON)
	public Response downloadAttachmentUsingJxl() throws IOException{
		File file = new File("./export.xls");
		try {
			if(file.exists()){
				file.delete();
			}
			file.createNewFile();
		} catch (IOException e) {
		} catch(Exception e){
		}
		
		excelService.createWorkbook(excelService.wriToExcel(getProductList()), file);
		ResponseBuilder response = Response.ok((Object) file);
		response.header("Content-type", "application/vnd.ms-excel");
	    response.header("Content-Disposition", "attachment; filename=excel.xls");
	    return response.build();
	}
	
	
	@Path("/poi/export")
	@GET
	@Produces (MediaType.APPLICATION_OCTET_STREAM)
	@Consumes(MediaType.APPLICATION_JSON)
	public Response downloadAttachmentUsingPOI() throws IOException{
		File file = new File("./export.xls");
		try {
			if(file.exists()){
				file.delete();
			}
			file.createNewFile();
		} catch (IOException e) {
		} catch(Exception e){
		}
		
		excelService.createWorkbookForPOI(excelService.wriToExcel(getProductList()), file);
		ResponseBuilder response = Response.ok((Object) file);
		response.header("Content-type", "application/vnd.ms-excel");
	    response.header("Content-Disposition", "attachment; filename=excel.xls");
	    return response.build();
	}
	
	private List<Product> getProductList(){
		List<Product> productList  = new ArrayList<Product>();
		productList.add(new Product("P115511", "IPod"));
		productList.add(new Product("P115511", "IPhone"));
		productList.add(new Product("P115511", "Teblate"));
		productList.add(new Product("P115511", "Laptop"));
		
		return productList;
	}
	
}
