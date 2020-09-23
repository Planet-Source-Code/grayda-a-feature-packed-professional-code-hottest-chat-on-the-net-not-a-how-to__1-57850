------------------------------------
NChat Alpha Build 10 Readme
------------------------------------

What is this file?

This file contains several sections, all dealing with the NChat Alpha Project. It is here to benifit people who own the source code, or the compiled application, who would like to know more about NChat and it's technical aspects. It is not a basic how-to guide for using the program. That is included with this source file.



Section 1: What is NChat?

NChat is a feature packed Chat client / server for Windows. It was first designed as a communications application for my school. Being the only person at school with a decent knowledge of Visual Basic (except of Nieumi of Solid Inc. :) ), this program took off and soon had a large following. Long after the NChat buzz died down at school, I decided to continue this project, developing it for others to use on home / work Networks.

NChat has HEAPS of features, with some that you won't find on other chat applications (Internet and Network), including:

> Over 60 Picture Smileys, ready to insert
> Customizable pictures next to your name
> FULL set of Administrator tools including:
   > Kick users
   > Re-Direct Users
   > Ban users by IP or username
   > Ghost users
   > Print User Info
> A fully working chat-bot (Called Notch, the NChat AutoBot) with:
   > Wildcard searching and matching
   > INI Based structure for easy programming
   > Multiple answers to a question ensure bot is never monotonous
   > Bot has idle chatter, determined after a certain interval
   > Bot can be updated with new content on-the-fly!
> Earn NCredits - NChat's very own currency!
> Spend NCredits in the store, and purchase some cool items!
> Full data logging, for people who want to learn more about NChat's data system
> WinMX Style chatroom "actions"
> Custom, funky green command buttons
> Load colour profiles to change NChat's appearance
> Private chat with unlimited people, with private Whiteboard!
> NChat supports Unlimited users, only limited by your bandwidth!
> Create and run your own chat room!
> Awesome Transparent PNG Splash screen
> Almost totally compatible with Windows 98 and ME!

> Plus HEAPS more features

NChat does not need one central server, because each copy of NChat runs on UDP Protocols, meaning that NChat is it's own server! You can run your own room without installing extra programs etc.



Section 2: What software does NChat require?

NChat requires the following files:

msvbvm60.dll <-- Standard Visual Basic File
oleaut32.dll <-- OLE Automation Library
olepro32.dll
asycfilt.dll 
stdole2.tlb  <-- OLE Automation
COMCAT.DLL 
msimg32.dll  <-- For Alpha Blending and other imaging stuff
GdiPlus.dll  <-- For Windows 2000 / XP ONLY. More imaging stuff for splash screen
scrrun.dll   <-- For File System Objects and other scripting stuff
RICHTX32.OCX <-- Standard Rich Text Box. A beefed up Text Box
comdlg32.ocx <-- Common Dialog control. Shows Open, Save, Print and colour boxes
MSCOMCTL.OCX <-- Common Windows Controls such as Progress Bars, Image Lists etc.
TABCTL32.OCX 
MSWINSCK.OCX <-- The heart of NChat. Lets us use Network Resources

If you do not have these files, then you can download the NChat Alpha setup file from:

http://www.solidinc.tk (Under the Downloads page, under the Applications category)

The setup file will install NChat Alpha Build 10, plus the required dependency files, or you can search on Google for the file names, or better yet, upgrade to Windows XP with VB6.0!


NChat has been sucessfully tested on the following Operating Systems:

Windows XP 	(100%)
Windows 2000 	(99%)
Windows ME	(80%)
Windows 98	(Not yet tested)
Windows 95	(Not yet tested)
Linux with Wine (0%, but tested with old version)

If you have tested NChat on any of the operating systems which haven't been tested (or even emulated OS's), then please forward the results to: firestorm_visual@hotmail.com to be included in the about box in the next major release. (Big accolades, I know :P)

At the moment, NChat has some major problems which makes it unsuitable for use in high-risk areas and all that other crap:

> The file transfers DO NOT work, and I can't work out why.
> When re-directing users to other rooms, the user is redirected up to 500 times. This is strange because no For loops, Do Loops or Timers are used when re-directing users.
> At times, data is not sent, or sent twice. I can't figure out why it does this. Must be a UDP problem. UDP has lots of troubles with that :)
> On Windows ME, text highlighting DOES NOT work. I think :)


Section 3: About Notch

Notch is NChat's AutoBot. He is an automatic room robot. He uses a simple form of "Simulated Intelligence". It's not AI, because Notch cannot learn from his experiences (yet)

Notch uses a structured INI file to interpret questions, and can be programmed to do almost anything you want him to do.

He can understand fragments of sentences (See the Notch help file, located in the "Help" folder in this directory), and can handle multiple questions in one line.

Notch also has "Idle Chatter", where if no talking is detected within a certain time frame, then he will speak a random phrase. This can be set from the "Admin Menu", under "Start / Stop Bot"


Section 4: What can we expect from NChat in the future?

> Notice how most network chats don't use Web-cams? What if the person is on a Wide Area Network? What if they need to show you something in another room? Yeeaaaaahhhh.... NChat may one day support Web-Cams

> The Notch file will be HEAVILY updated some day, so he isn't as stupid. If you can't wait, then feel free to do this for me! E-Mail the results to: firestorm_visual@hotmail.com

> One day I'll get around to commenting the WHOLE project, so it's easier to understand

> Testing will commence on Windows 98 and lower, to ensure it actually works. 

> All my dirty little code fixes will be removed, streamlining the code significantly (Check out the Text sub. 2 lines of code are used to determine if smileys are to be checked. Can make it one or even none (using booleans))

> I'll sort out the rest of the bugs later

> And these documents... Starting to mess up the main folder :(