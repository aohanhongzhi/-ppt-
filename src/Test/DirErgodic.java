package Test;

import java.io.File;
import java.io.IOException;
 
public class DirErgodic {
 
	private static int depth=1;
	
	public static void find(String pathName,int depth) throws IOException{
		int filecount=0;
		//获取pathName的File对象
		File dirFile = new File(pathName);
		//判断该文件或目录是否存在，不存在时在控制台输出提醒
		if (!dirFile.exists()) {
			System.out.println("do not exit");
			return ;
		}
		//判断如果不是一个目录，就判断是不是一个文件，时文件则输出文件路径
		if (!dirFile.isDirectory()) {
			if (dirFile.isFile()) {
				System.out.println(dirFile.getCanonicalFile());
			}
			return ;
		}
		
		for (int j = 0; j < depth; j++) {
			System.out.print("  ");
		}
		System.out.print("|--");
		System.out.println(dirFile.getName());
		//获取此目录下的所有文件名与目录名
		String[] fileList = dirFile.list();
		int currentDepth=depth+1;
		for (int i = 0; i < fileList.length; i++) {
			//遍历文件目录
			String string = fileList[i];
			//File("documentName","fileName")是File的另一个构造器
			File file = new File(dirFile.getPath(),string);
			String name = file.getName();
			//如果是一个目录，搜索深度depth++，输出目录名后，进行递归
			if (file.isDirectory()) {
				//递归
				find(file.getCanonicalPath(),currentDepth);
			}else{
				//如果是文件，则直接输出文件名
				for (int j = 0; j < currentDepth; j++) {
					System.out.print("   ");
				}
				System.out.print("|--");
				System.out.println(name);
				
			}
		}
	}
	
	public static void main(String[] args) throws IOException{
		find("D:\\", depth);
	}
}
