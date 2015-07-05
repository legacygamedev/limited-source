/**
 * This class is designed to be packaged with a COM DLL output format.
 * The class has no standard entry points, other than the constructor.
 * Public methods will be exposed as methods on the default COM interface.
 * @com.register ( clsid=E5E3960F-01E5-44C7-A207-8F7E5631CD0F, typelib=84E7D8D8-1843-4949-9BF7-DF15B7B2D9A7 )
 */

import com.ms.com.*;

/**
 * @com.register ( clsid=C56BA553-5D28-4A1C-B05B-ED53C87AFA33, typelib=D51A728E-F6B7-48DC-AF72-4CF7BE535AB5 )
 */
public class plugin
{
	private Win32 fred;
	private Dispatch pDisp;
	
	public void SetHost(Object newref){ 	
		
		Variant intMenuId = new Variant(0) ; 
		Variant intStartupArg = new Variant(1) ; 
		
		/*
		//just playing around..you could also get the dispID this way and use that
		String[] dispNames = { "RegisterPlugin" };
		int[] dispIds;
		dispIds = pDisp.getIDsOfNames(newref,null,0,dispNames);
		*/
		
		//for each plugin you want to register you will have to fill out
		//the myArgs array and call the invoke method below, then handle
		//each arg in the StartUp sub below.
		Object[] myArgs = { intMenuId, "J++ Sample Plugin", intStartupArg };
	
		pDisp.invoke(newref,null,"RegisterPlugin",0,0, Dispatch.Method ,myArgs,null);		
	}
	
	public void StartUp(int myArg){
		String msg = new String("J++ Plugin Sample being called with Arg=");
		msg+=myArg;
		fred.MessageBox(0,msg,"J++ Sample Plugin",0); 
	}
	
}
