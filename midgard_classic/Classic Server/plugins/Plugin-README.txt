Plugins are experimental right now so dont expect so much from them.
I have made and included the code to two sample plugins

1. Trivia Bot.
2. Swear Kick Bot.

The trivia bot loads a list of questions from questions.txt and then
proceeds to host a game of trivia on the map it is on.

The swear kick bot has a array in the code holding some sample "curse"
words that if detected in global chat or the map its on it will kick the
person who said it. The bot does not beed to be admined to kick the person.


Plugins are preety much bots that can be programmed in Microsoft Visual
Basic 6.0 As a ActiveX DLL. The two plugins i have included can be edited.
I really dont care what you do with them preety much since there used to
show how the coding works at the moment. I will be setting up a online 
guide with the (Codes.) list and the proper way to use them. And more sample
code once i get some time.


Plugins are preety simple to use. You must have Microsoft Visual Basic 6.0
to compile them. If not get a friend to compile them and follow the steps
after compiling. Also on a side note you can only use one plugin at a time.
due to a bug i will be fixing hopefully in Uber RPG AlphaR3.

-Compiling
To compile the plugin open its source code. and then goto File>Make
Pluginname.dll Pluginname as the name of the plugin. Save it in the plugins
folder that contains dictionary.txt and regsvr32.exe You also need to create
a account for the plugins to use. You can change the name/pass it uses in the
modMain module where Username = "the username you are going to use" and
password = "the password you want"

-Registering
This is a simple process. go into the plugins folder and click+drag the .dll
file you just made onto the regsvr32.exe If successfull it should say
DllServerRegistered "the plugin path.dll" successfully. Or something close at
least. Then you can delete the .exp and .lib files that Visual Basic created.

-Using
After you follow the steps above using the plugin is simple. You run the
server then click options on the tab menu. And then click Open plugin manager.
Then select the plugin from the list on the left and click Run Selected Plugin
Hopefully if you did this correctly a new window will appear. If it lags a
for a few seconds dont worry the lag will go away. For the trivia bot click
the Run Trivia button to start the game of trivia. The swear kick bot will
start working as soon as you run it and the window shows up. You can only use
one plugin at a time for now. Hopefully in R3 i will have that worked out so
you can run 50 of them if you wanted. But combining the swear kick bot and the
trivia bot would be about 5 whole seconds of work.



-ADDITIONAL COMMENTS & SUPPORT-
If you want more information or help on plugins visit the forums. for a direct
topic URL copy and paste the url below into your favorite web browser and hit
go
http://uber.pcbot2k.com/forum/viewforum.php?f=21

-Pc
