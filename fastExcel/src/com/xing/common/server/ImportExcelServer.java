/**
 * 
 */
package com.xing.common.server;

import java.io.File;
import java.util.List;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 上午10:02:19   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public interface ImportExcelServer {
	
	public <T> List<T> importExcelServer(File excelFile, Object reflectEntity, Class<T> dataEntity, int rowStartNum)throws Exception;

}
