package Test;

import java.io.File;
import java.io.IOException;
 
public class DirErgodic {
 
	private static int depth=1;
	
	public static void find(String pathName,int depth) throws IOException{
		int filecount=0;
		//��ȡpathName��File����
		File dirFile = new File(pathName);
		//�жϸ��ļ���Ŀ¼�Ƿ���ڣ�������ʱ�ڿ���̨�������
		if (!dirFile.exists()) {
			System.out.println("do not exit");
			return ;
		}
		//�ж��������һ��Ŀ¼�����ж��ǲ���һ���ļ���ʱ�ļ�������ļ�·��
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
		//��ȡ��Ŀ¼�µ������ļ�����Ŀ¼��
		String[] fileList = dirFile.list();
		int currentDepth=depth+1;
		for (int i = 0; i < fileList.length; i++) {
			//�����ļ�Ŀ¼
			String string = fileList[i];
			//File("documentName","fileName")��File����һ��������
			File file = new File(dirFile.getPath(),string);
			String name = file.getName();
			//�����һ��Ŀ¼���������depth++�����Ŀ¼���󣬽��еݹ�
			if (file.isDirectory()) {
				//�ݹ�
				find(file.getCanonicalPath(),currentDepth);
			}else{
				//������ļ�����ֱ������ļ���
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
