# LookOut (Fix Version)<img src="src/chrome/skin/LOicon-64.png" width="80" align="left">

[This is a copy of the source code for the addon by Oleksandr](https://addons.thunderbird.net/nn-NO/thunderbird/addon/lookout-fix-version/)

LookOut is a plugin which allows Thunderbird to interface with Microsoft’s mail tools by decoding metadata and attachments encapsulated/embedded in a [TNEF (Transport Neutral Encapsulation Format)](https://en.wikipedia.org/wiki/Transport_Neutral_Encapsulation_Format) encoded attachment (aka winmail.dat).

Support: Thunderbird and SeaMonkey


## About this Add-on
Original LookOut add-on by Aron Rubin for Mozilla Thunderbird is abandoned September 30, 2011:
https://addons.mozilla.org/addon/lookout/

Does it mean its death? Of course not! It has been reborn! I brought it back to life!
Now it is called LookOut (fix version) and works with all new Thunderbird versions.

LookOut (fix version) - based on last version LookOut+ 1.2.15 by Attila K. Mergl (Last Updated: October 15, 2013)
https://addons.mozilla.org/addon/lookout-1/

Please before use, remove or disable the original LookOut and LookOut+ !

## What is LookOut?

LookOut is a plugin which allows Thunderbird (and SeaMonkey) to interface with Microsoft’s mail tools by decoding metadata and attachments encapsulated/embedded in a TNEF encoded attachment (aka winmail.dat).
The TNEF decoding engine was inspired (with permission) by the [TNEF project](https://github.com/verdammelt/tnef). LookOut has a full MAPI decoder and the ability to transcode some of those properties to [iCalendar (ics)](https://en.wikipedia.org/wiki/ICalendar) and [vCard (vcf)](https://en.wikipedia.org/wiki/VCard) formats.
As an option all [MAPI](https://en.wikipedia.org/wiki/MAPI) properties within a email can be viewed so other types of information can be deciphered.

# Debugging

Debugging can be enabled in the settings menu. Output will be presented in the error console.

Note: Re-opening a message is necessary to see any debug outputs after setting change.
