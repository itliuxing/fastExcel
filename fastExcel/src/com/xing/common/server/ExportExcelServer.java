/**
 * 
 */
package com.xing.common.server;

import java.io.File;
import java.util.List;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 下午04:06:01   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public interface ExportExcelServer {
	
	<T> String exportExcelServer( List<T> dataList , Object reflectEntity, File excelFile, int startRomNum )throws Exception ;
	
}
