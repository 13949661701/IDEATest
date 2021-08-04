package com.sqp.utils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.lang.StringUtils;

/**
 * 字符串工具类
 * @author crfchina
 * 上海信而富企业管理有限公司
 */
public class StringUtil {
	
	private static final String ROWKEY_SPLIT_SYMBOL = "_";
	//  少数民族姓名中符号
	private static final String ETHNIC_MINORITY_NAMES_SYMBOL = "·";
	
	//  含字母正则
	private static final String LETTER_REX =  "[A-Za-z]";
	//  父母昵称正则
	private static final String ELDER_NAME_REX =  "爸|父|妈|母";
	//  平辈昵称正则
	private static final String EQUALS_NAME_REX =  "兄|哥|弟|姐|妹";
	//  以阿开头的姓名正则
	private static final String AH_NAME_REX =  "^阿.*$";
	
	private static final Pattern LETTER_PATTERN = Pattern.compile(LETTER_REX);
	private static final Pattern ELDER_NAME_PATTERN = Pattern.compile(ELDER_NAME_REX);
	private static final Pattern EQUALS_NAME_PATTERN = Pattern.compile(EQUALS_NAME_REX);
	private static final Pattern AH_NAME_PATTERN = Pattern.compile(AH_NAME_REX);
	

	/**
	 * 根据传入phone计算对应的报名的末位数字
	 * @param str
	 * @return
	 */
	public static int getRowKeyPreNum(String phone) {
		int code = phone.hashCode();
		return Math.abs(code)  % 10;
	}
	
	/**
	 * 获取rowKey
	 * @param phone
	 * @return
	 */
	public static String getPhoneBookRowKey(String phone) {
		return getRowKeyPreNum(phone) + phone;
	}
	
	/**
	 * 获取rowKey
	 * @param phone
	 * @return
	 */
	public static String getIpRowKey(String ip,String date) {
		return getRowKeyPreNum(ip) + ROWKEY_SPLIT_SYMBOL + ip + ROWKEY_SPLIT_SYMBOL + date;
	}
	
	/**
	 * hdfs文件验证字符是否是空
	 * @param str
	 * @return
	 */
	public static boolean isBlank(String str) {
		return StringUtils.isBlank(str) || "null".equals(str) || "(null)".equals(str) || "\\N".equals(str);
	}
	
	
	public static boolean isNotBlank(String str) {
		return StringUtils.isNotBlank(str) && !"null".equals(str) && !"(null)".equals(str) && !"\\N".equals(str);
	}
	
	/**
	 * 判断一个字符串是否包含字母
	 * @param str
	 * @return
	 */
	public static boolean containsLetter(String str) {
		return LETTER_PATTERN.matcher(str).find();
	}
	
	/**
	 * 判断一个字符串是否长辈昵称字符
	 * @param str
	 * @return
	 */
	public static boolean containsElderName(String str) {
		return ELDER_NAME_PATTERN.matcher(str).find();
	}
	
	/**
	 * 判断一个字符串是否以阿开头的姓名
	 * @param str
	 * @return
	 */
	public static boolean startWithAhName(String str) {
		return AH_NAME_PATTERN.matcher(str).find();
	}
	
	/**
	 * 判断一个字符串是否包含平辈昵称字符
	 * @param str
	 * @return
	 */
	public static boolean containsEqualsName(String str) {
		return EQUALS_NAME_PATTERN.matcher(str).find();
	}
	
	/**
	 * 根据姓名获得其姓氏
	 * @param name
	 * @return
	 */
	public static String getSurname (String name) {
		if (StringUtil.isBlank(name)) {
			return null;
		} else {
			int nameLength = name.length();
			//  是否是少数民族姓名
			if (name.contains(ETHNIC_MINORITY_NAMES_SYMBOL)) {
				return name.split(ETHNIC_MINORITY_NAMES_SYMBOL)[0];
			} else {
				//  名字长度少于4位取第一个字符，反之取俩个
				if (nameLength < 4) {
					return name.substring(0, 1);
				} else {
					return name.substring(0, 2);
				}
			}
		}
	}
	
	/**
	 * 过滤去除字符串中的空格、回车、换行符、制表符
	 * 
	 * @param strValue
	 * @return
	 */
	public static String filter(String strValue) {
		String dest = "";
		try {
			if (StringUtils.isEmpty(strValue)) {
				return dest;
			}
			if (StringUtils.isEmpty(strValue.trim())) {
				return dest;
			}
			// 过滤
			Pattern p = Pattern.compile("\\s*|\t|\r|\n");
			Matcher m = p.matcher(strValue.trim());
			dest = m.replaceAll("");
			// 去除其中占四个字节的特殊汉字、表情等
			if (StringUtils.isNotBlank(dest)) {
				dest = dest.replaceAll("[\\ud800\\udc00-\\udbff\\udfff\\ud800-\\udfff]", "");
			}
			return dest;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return dest;
	}
	
    /**
     * 验证是否是数字
     * @param str
     * @return
     */
	public static boolean isDecimal(String str) {
		return Pattern.compile("([1-9]+[0-9]*|0)(\\.[\\d]+)?").matcher(str).matches();
	}
}
