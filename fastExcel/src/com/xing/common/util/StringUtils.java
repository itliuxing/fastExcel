/**
 * 
 */
package com.xing.common.util;

/**
 * 类描述：   		
 * 创建时间：	2013-9-27 上午10:36:16   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public class StringUtils {
	
/*	private static StringUtils stringUtils ;
	private StringUtils(){};
	
	public synchronized static StringUtils getInstance(){
		if( stringUtils == null ){
			stringUtils = new StringUtils() ;
		}
		return stringUtils ;
	}*/
	
	public static boolean isNoEmpty( String value ){
		boolean booleanValue = true ;
		if( value == null ){
			booleanValue = false ;
		}else if( value.length() == 0 ){
			booleanValue = false ;
		}
		return booleanValue ;
	}
}
