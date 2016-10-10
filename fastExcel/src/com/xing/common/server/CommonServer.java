/**
 * 
 */
package com.xing.common.server;

import java.io.File;
import java.util.Map;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 上午10:15:07   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public interface CommonServer {
	
	public Map<String,Integer> headHandlerServer( File excelFile )throws Exception ;
	
	public Map<String, Object> reflectEntityHandlerServer( Object reflectEntity )throws Exception;
	
	public Map<String,Integer> twoLevelHandlerServer( Map<String,Integer> excelMap , Map<String,Object> reflectMap );

}
