# Flight Sim Tool v0.6.0

## About
This little app uses SimConnect to continuously read position data from MSFS2020.  Meh I hear you say, so what - lots of far more capable apps out there that'll do that.  Ahh true but despite lots of searching I have not yet found one that will allow me to instantly slew my aircraft to any position on the globe.  Hence this project was born.

Update for v0.5.0 - The above is no longer true, there are now many utilities that provide slewing functions available but that is no longer the primary purpose for this app.  However I have left it in because I still use it in conjunction with one of my favourite FS companion apps - LittlNavMap.

Took me a while to get to grips with SimConnect but thanks to these two posts 
https://www.fsdeveloper.com/forum/threads/msfs-2020-managed-simconnect.448446/#post-855955
https://forums.flightsimulator.com/t/a-minimalistic-simconnect-example/261878
and the FS2020 SDK documentation, I was able to understand enough to get the functionality I wanted.

Central Repository: https://github.com/Elephant42/FS_Tool


## Installation
Installation is really simple. You only have to download the zip file and then un-zip it anywhere you like.  Double click on FS_Tool.exe to run the app.  It will sit there waiting for a connection to the sim and show the connected status in the title bar.  The "Slew" button will not be enabled until the app is connected to the sim.


## How To Use
Pretty simple really - tell it a geographic coordinate and push the "Slew" button.

There are two ways to give it a position.
1. Type it in to the two text boxes, Latitude and Longitude.
2. Copy a position to the clipboard and then click the "Copy From Clipboard" button.

As long as it's able to parse the position string the "Slew" button will be enabled - click it and voila.

The position parser currently understands the format from Google maps as well as all the formats from LittleNavMap.

It's pretty cool - right click anywhere on the LNM map and select "More->Copy nnnn yyyy to clipboard", then click the "Copy From Clipboard" button in this app followed by "Slew" and watch the magic.

All the other parameters in the plane remain at their pre-slew values - speed, height and heading.


## Bonus Function
There is also a secondary function available which will allow you to change the simulation rate from x1 to x8.  This can be done from the main form or from the Tray Icon context menu.


## Big Bonus Function :)
As of version 0.5.0 the app is now able to map HID events, specifically for joysticks, to SimConnect events and to keystrokes.  This can be used with any HID joystick but is primarily aimed at the Honeycomb Bravo throttle quadrant.


## Tray Icon
The app is meant to run in the system tray and has a context menu to suit.  While minimised you can use the Tray Icon context menu to set the required Sim Rate, slew to the current contents of the clipboard and activate joystick event mapping.  There is also a settings option which will make the app go straight to the tray as soon as it's launched.


## Building
The app is a bog standard VB.Net project and should build on any version of Visual Studio later than and including VS2017.

The app utilises code from the HIDSharp library, Copyright 2010-2019 James F. Bellinger <http://www.zer7.com/software/hidsharp>

One caveat is that you will need the MSFS2020 SDK installed in order for the SimConnect libraries to be found.  The project is configured to look for them in "C:\MSFS SDK" so if your SDK is installed to a different folder you will need to remove and re-add the Microsoft.FlightSimulator.SimConnect reference and change the following post-build events to point to the correct path:
xcopy "C:\MSFS SDK\SimConnect SDK\lib\SimConnect.dll" "$(TargetDir)" /y
xcopy "C:\MSFS SDK\Samples\SimvarWatcher\SimConnect.cfg" "$(TargetDir)" /y


## License
My code is licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007
The included HIDSharp library code which I utilise for reading joystick events is licensed under the Apache License, Version 2.0.


## Changelog

## [0.1.0]
### Initial Commit
- **[Slew]** Ability to slew aircraft to any position on the globe.

## [0.5.0]
### 
- **[SimConnect]** All functions separated out to a library DLL.
- **[SimConnect]** Able to inject low level keystroke events.
- **[Joysticks]** Able map HID joystick events to SimConnect events and/or keystroke events.

## [0.6.0]
### Minor Enhancements
- **[Joysticks]** Added long press events.
- **[Joysticks]** Added ability to send multiple events with a single HID Event.
- **[Joysticks]** Added ability to specify the data to send with each SimEvent.
- **[Joysticks]** Added ability to send an event multiple times in a single HID event.
- **[Joysticks]** Added ability to specify both press and release events for a HID event.
- **[General]** Documentation updates and improvements.


## Known Issues - None.


## Future Plans
I consider this little project to be a WIP and plan to add further functionality as my expertise with SimConnect grows.  I welcome contributions from anyone with said expertise, that they care to make.  One feature I would dearly love to implement is a moving map which can be used for direct click slewing.  Need to start digging into Google and other maps APIs...


## FAQ

**Q: What does it all mean?**

A: It's all meaningless and we are all going to die. https://www.youtube.com/user/erfmufn
