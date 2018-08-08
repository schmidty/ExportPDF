
public class NameCase {

	public static String transformCase(String s) //{{{
	{
		s = s.trim();
		StringBuffer sb = new StringBuffer();
		boolean wasSpace = true;
		boolean wasSpaceChar = false;
		boolean wasLowerCase = false;
		int cnt = s.length();
		int charCnt = 0;

		for(int i=0; i<cnt; i++) {
			char c = s.charAt(i);
			System.out.print("Char: " + c);

			// This has to be more than 2 characters
			if(c=='-' || c=='\'') {
				System.out.println(" --> Single quote-hyphen");
				wasSpace = true;
				wasSpaceChar = false;
				wasLowerCase = false;
				sb.append(c);
			} else 
			if(c==' ') {
				charCnt = 0;
				System.out.println(" --> Space");
				if(!wasSpaceChar) {
					wasSpace = true;
					wasSpaceChar = true;
					wasLowerCase = false;
					sb.append(c);
				} 				
			} else {
				charCnt++;
				System.out.println(" --> " + charCnt + " Char");
				if(wasSpace) {
					sb.append(Character.toUpperCase(c));
					wasLowerCase = false;
				} else
				if( wasSpace && charCnt == 2) {
					sb.append(Character.toUpperCase(c));
				}
				else {
					// Count number of non-space characters
					if((wasLowerCase == true) && (Character.toUpperCase(c)==c)) {
						sb.append(Character.toUpperCase(c));
						wasLowerCase = false;
					} else {
						sb.append(Character.toLowerCase(c));
						wasLowerCase = (Character.toLowerCase(c) == c);
					}
				}
				wasSpace = false;
				wasSpaceChar = false;
			}
		}

		String ret = sb.toString();
		if(!s.equals(ret)) {
			System.out.println("NameCase: " + s + " --> " + ret);
		}
		return ret;
	}//}}}

	public static void main(String argv[]) //{{{
	{
		System.out.println(transformCase(argv[0]));
	}//}}}

}
