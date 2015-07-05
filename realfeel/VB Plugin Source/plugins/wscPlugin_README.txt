
--> regsvr32 the wsc file before it will work, 

This will show you how you can easily make plugins from the 
comfort of your favorite scripting language.

Basically any language that has a windows scripting engine, can
be used to make your very own plugin.

This includes PERLScript, PScript, Python as well as the standard
JScript, VBScript.

So how does this all work you ask?

The great guys at Microsoft have this cool thing call Windows Script
COmponents. Basically its an XML script file that the scripting
runtime can use as an on the fly COM object :)

The wscPlugin_2.wsc file included in this package is a sample
script written in VBScript that is an actual working plugin for
this framework by implementing the plugin class.

This wsc file was created with the Microsoft Script Component Wizard.

You can download it here:

http://www.microsoft.com/downloads/details.aspx?FamilyId=408024ED-FAAD-4835-8E68-773CCC951A6B&displaylang=en

A couple notes on creating your own components and getting this one to work.

1) If you just modify this one make sure to change one of the numbers in the clsid
   so that it does not conflict with other ones you make (if you make more than one)

2) make sure that you end your progid = xxxxxx.plugin
   The classname has to be .plugin

3) this may sound a little crazy, but you have to register the wsc file with
   regsvr32 to make sure the registry is setup right for it to work.
   to do this,
  
   goto start button -> run 
   type regsvr32 
   now drag and drop the wsc file from explorer into the textbox and its path will show up
   hit ok and it should give you a messagebox that it was registered ok. 

   note: using the wizard to create a new one didnt seem to register it, i still had to do
         it manually

   note2: once you have registered the file, you cant move it, if you do..you will have to 
          reregister it to update its location in the registry

   note3: this is really cool technique :)

   note4: if you wanted to make an propiratary plugin using this technique, you could use the
           script encoder and use vbe or jse for the script lanaguage to keep others away from 
           the code...although I think somewhere people did figure out how to crack that.

   note5: one thing that is kind of annoying is debugging, once plugins are loaded, the host already
          has the object in memory, so changes to the script file mean you have to reload the plugin
	  . If your framework were to support unloading/loading of plugins it could help.

	  really though, since these are just script commands, you can usally debug them for the most
          part through other means (such as using a script control on a form etc...)





