package util;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DigitUtil {

	/**
	 * 截取出字符串中第一个连续的数字
	 * 
	 * @param 字符串
	 * @return 字符串中第一个连续的数字
	 */
    public static Integer getNumbers(String content) {  
        Pattern pattern = Pattern.compile("\\d+");  
        Matcher matcher = pattern.matcher(content);  
        while (matcher.find()) {  
           return Integer.parseInt(matcher.group(0));  
        }  
        return null;  
    }  
    
    /**
     * 将大写字母转换为数字，A->1, B->2,余类推
     * 
     * @param letter,输入的大写字母
     * @return 大写字母对应的数字 
     */
	public static int charToNum(char letter){
		return letter - 64;
	}
	
}