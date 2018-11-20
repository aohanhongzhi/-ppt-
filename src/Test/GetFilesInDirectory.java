package Test;

import java.io.File;
import java.util.ArrayList;

public class GetFilesInDirectory {

	public static void main(String[] args) {
		ArrayList<String> al = new ArrayList<String>();
		String pathName="/";
		File dirFile = new File(pathName);
		//判断该文件或目录是否存在，不存在时在控制台输出提醒
		if (!dirFile.exists()) {
			System.out.println("do not exit");
			return ;
		}
		String[] fileList = dirFile.list();
		for (int i = 0; i < fileList.length; i++) {
			String string = fileList[i];
			//File("documentName","fileName")是File的另一个构造器
			File file = new File(dirFile.getPath(),string);
			String path = file.getAbsolutePath();
			System.out.println(path);
			if(path.endsWith("ppt") || path.endsWith("PPT") || path.endsWith("pptx") || path.endsWith("PPTX")){
				al.add(path);
			}
		}
		
		// ppt 2 pdf 
		
		
	}
}
