package com.blackmcat.excel;

public class ExcelUtil {
	/**
	 * 设置合并标题单元格的长度
	 * 
	 * */
	public static String mergeTitle(int mergeLen){
		String  merges = transverseCellMerge("$A$1",mergeLen);
		return merges;
	}
	
	/**
	 * 横向合并单元格
	 * @param cell 起始单元格位置
	 * @param size 合并单元格的数量
	 * */
	public static String transverseCellMerge(String cell,int size){
		String s= cell.substring(1, 2);
		int loopSize = size/26;
		int cellLen = size%26;
		String z="";
		//1-26 A-Z 26-52 AA-AZ ...
		if(loopSize ==0){
			byte a = (byte) s.charAt(0);
			byte b = (byte) (a+cellLen-1);
			char c = (char) b;
			z=""+c;
		}else{
			byte a = (byte) s.charAt(0);
			byte b = (byte) (a+cellLen-1);
			char c = (char) b;
			char d =(char) (byte) (64+loopSize);
			z=""+d+""+c;
		}
		String  merges = cell+":$"+z+"$"+cell.substring(cell.length()-1, cell.length());
		return merges;
	}
	/**
	 * 竖向合并单元格
	 * @param cell 起始单元格位置
	 * @param size 合并单元格的数量
	 * */
	public static String verticalCellMerge(String cell,int size){
		String s= cell.substring(1, 2);
		int ss  = Integer.parseInt(cell.substring(cell.length()-1, cell.length()))+size-1;
		String  merges = cell+":$"+s+"$"+ss;
		return merges;
	}
	
	
	public static void main(String[] args) {
		System.out.println(mergeTitle(25));
	}
	
}
