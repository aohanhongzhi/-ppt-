package Test;

import java.util.*;

public class IteratorTest {
	public static void main (String args[]) {
		
		List<Object> l = new ArrayList<Object>();
		 l.add("aa");
		// l.add("bb");
		// l.add("cc");
		 for (Iterator<Object> iter = l.iterator(); iter.hasNext();) {
		     String str = (String)iter.next();
		     System.out.println(str);
		 }
		 /*迭代器用于while循环
		 Iterator iter = l.iterator();
		 while(iter.hasNext()){
		     String str = (String) iter.next();
		     System.out.println(str);
		 }
		 */
		
	}
}
