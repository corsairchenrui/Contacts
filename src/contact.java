import java.util.HashSet;
import java.util.Set;

public class contact {
	protected String lname;
	protected String fname;
	protected String mobile;
	protected String tel;
	protected Set<String> category = new HashSet<String>();
	@Override
	public String toString(){
		StringBuffer ret = new StringBuffer();
		ret.append("BEGIN:VCARD\n");
		ret.append("VERSION:3.0\n");
		ret.append("N:;"+fname+lname+";;;\n");
		ret.append("FN:"+fname+lname+"\n");
		if(mobile!=null&&!mobile.isEmpty())
			ret.append("TEL;TYPE=CELL:"+mobile+"\n");
		if(tel!=null&&!tel.isEmpty())
			ret.append("TEL;TYPE=WORK:"+tel+"\n");
		for(String enumCatogary:category)
			ret.append("CATEGORIES:"+enumCatogary+"\n");
		ret.append("X-WDJ-STARRED:0\nEND:VCARD\n");
		return ret.toString();
	}
	public boolean equals(contact another) {
		if (this.lname.equals(another.lname)
				&& this.fname.equals(another.fname)) {
			if (this.mobile!=null&&another.mobile!=null&&this.mobile.equals(another.mobile))
				return true;
			if(this.tel!=null&&another.tel!=null&&this.tel.equals(another.tel))
				return true;
		}
		return false;
	}
}