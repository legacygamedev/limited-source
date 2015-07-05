public class Win32
{
	/**
	 * @dll.import("USER32", auto) 
	 */
	public static native int MessageBox(int hWnd, String lpText, String lpCaption, int uType);
}
