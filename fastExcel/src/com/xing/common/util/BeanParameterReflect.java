/**
 * 
 */
package com.xing.common.util;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 类描述：   		将传入进来的对象，通过反射以map(name:value),方式解析出对象内的参数值
 * 创建时间：	2013-8-30 下午02:06:57   
 * 创建人：		LiuXing   
 * 修改备注：   	
 * @version		V1.0 
 */
public class BeanParameterReflect {
	
	/***
	 * 将java对象转换为map
	 * @param javaBean			传入对象
	 * @return
	 * @throws Exception		发生异常
	 */
	public static Map<String, Object> reflectPrarameter(Object javaBean)throws Exception{
		Map<String, Object> map = new HashMap<String, Object>() ;
		
		//获取对象的全部属性，以及值 然后封装成map
		Field[] fields = javaBean.getClass().getDeclaredFields() ;
		if( fields.length > 0 ){
			for( int i=0 ; i<fields.length ; i++ ){
				fields[i].setAccessible(true) ;
				map.put( fields[i].getName() , fields[i].get(javaBean) ) ;
			}
		}
		return map ;
	}
	
	/****
	 * 自动匹配关于两个相同对象的相同属性的值是否相等，不相等的话.返回一个字符串数组：属性,第一个对象的值,第二个对象的值
	 * @param javaBean		第一个对象
	 * @param targetBean	第二个对象
	 * @return
	 * @throws Exception
	 */
	public static List<String[]> reflectPrarameter(Object javaBean,Object targetBean)throws Exception{
		List<String[]> dataSource = new ArrayList<String[]>() ;
		//获取两个对象的全部属性，以及值   然后判断这个相同属性的值是否相同，不同的时候就需要记录，被修改
		Field[] fields = javaBean.getClass().getDeclaredFields() ;
		Field[] targetFields = javaBean.getClass().getDeclaredFields() ;
		if( fields.length > 0 && targetFields.length > 0){	//对象的属性长度不能为0
			for( int i=0 ; i<fields.length ; i++ ){
				fields[i].setAccessible(true) ;
				for( int targetFieldsi=0 ; i<targetFields.length ; targetFieldsi++ ){
					targetFields[targetFieldsi].setAccessible(true) ;
					if( targetFields[targetFieldsi].getName().equals( fields[i].getName() ) ){
						//判断两个对象的属性相同时，再判断属性值是否相同。不相同的话需要记录下，并传出做修改
						if( ! fields[i].get(javaBean).equals( targetFields[targetFieldsi].get(targetBean) ) ){
							dataSource.add(new String[]{ fields[i].getName() , fields[i].get(javaBean).toString() , targetFields[targetFieldsi].get(targetBean).toString() }) ;
						}
					}
				}
			}
		}
		return dataSource ;
	}

}
