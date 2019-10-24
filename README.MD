# NOTICE: New maintainer wanted

# LookOut (Fix Version)<img src="src/chrome/skin/LOicon-64.png" width="80" align="left">

**[Please Read the F.A.Q.](https://github.com/TB-throwback/LookOut-fix-version/wiki/F.A.Q.)**

[This is a copy of the source code for the add-on by Oleksandr](https://addons.thunderbird.net/addon/lookout-fix-version/)

LookOut is a plugin which allows [Thunderbird](https://www.thunderbird.net/) to interface with Microsoft’s mail tools by decoding metadata and attachments encapsulated/embedded in a [TNEF (Transport Neutral Encapsulation Format)](https://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format) encoded attachment (aka winmail.dat).

Support: [Thunderbird](https://www.thunderbird.net/)

**For legacy support please see the [Legacy Branch](https://github.com/TB-throwback/LookOut-fix-version/tree/Legacy)**

## New Contributors are Welcome

LookOut is now over 7 years old and is in need of a lot of work. If you feel you have anything to contribute please feel welcome to open a pull request or an issue.

## About this Add-on
Original LookOut add-on by Aron Rubin for Mozilla Thunderbird was abandoned September 30, 2011:
https://addons.thunderbird.net/addon/lookout/

Does it mean the add-on’s death? Of course not! It has been reborn! I brought it back to life!
Now it is called LookOut (fix version) and works with all new Thunderbird versions.

LookOut (fix version) is based on the last version of LookOut+ 1.2.15 by Attila K. Mergl (Last Updated: October 15, 2013)
https://addons.thunderbird.net/addon/lookout-1/

Before use, please remove or disable the original LookOut and LookOut+ !

## What is LookOut?

LookOut is a plugin which allows [Thunderbird](https://www.thunderbird.net/) (and [SeaMonkey](http://www.seamonkey-project.org/)) to interface with Microsoft’s mail tools by decoding metadata and attachments encapsulated/embedded in a TNEF encoded attachment (aka winmail.dat).
The TNEF decoding engine was inspired (with permission) by the [TNEF project](https://github.com/verdammelt/tnef). LookOut has a full [MAPI](https://en.wikipedia.org/wiki/MAPI) decoder and the ability to transcode some of those properties to [iCalendar (ics)](https://en.wikipedia.org/wiki/ICalendar) and [vCard (vcf)](https://en.wikipedia.org/wiki/VCard) formats.
As an option all MAPI properties within a email can be viewed so other types of information can be deciphered.

# Debugging

Debugging can be enabled in the Settings menu. Output will be presented in the Error Console.

Note: Re-opening a message is necessary to see any debug outputs after setting change.
