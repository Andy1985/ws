class Result 
{
	private byte vType;
	private int vLen;
	
	public void setValue(byte type,int len)
	{
		this.vType = type;
		this.vLen = len;
	}

	public byte getType()
	{
		return this.vType;
	}

	public int getLen()
	{
		return this.vLen;
	}
}	
