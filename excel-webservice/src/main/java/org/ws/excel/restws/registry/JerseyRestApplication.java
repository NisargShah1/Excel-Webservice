/**
 * @author {Nisarg Shah}
 * 21-Jun-2015
 * com.infosys.barclays.bli.restws.registry JerseyRestApplication.java
 */
package org.ws.excel.restws.registry;

import org.glassfish.jersey.server.ResourceConfig;
import org.ws.excel.restws.ws.ExcelExportWs;


/**
 * Jersey Restful Services configuration
 * @author NISARG
 *
 */
public class JerseyRestApplication extends ResourceConfig
{
	
	/**
	 * Register JAX-RS application components.
	 */
	public JerseyRestApplication(){
		register(ExcelExportWs.class);
	}
}
