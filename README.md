# GamescareRT4K
Switching app for Gamescare SCART Switch and the Retrotink4K.

A GamesCare SCART Switch on the local network can be accessed via http://gscartsw.local (by default) - this script will allow you to select the input ports directly instead of using the WebUI, as part of doing so, the relevant "Profile" button will be simulated for an attached RetroTink4K scaler device (https://www.retrotink.com/product-page/retrotink-4k).  Some additional buttons are also added for direct control of a few Rt4K functions.

Based on the initial script from Carter300: https://github.com/carter300/RetroTink4K-PCRemote/tree/main

The Retrotink must be connected to the PC via USB-C cable and recognised in the device manager.  This script does not _Currently_ support serial-over-DSUB, that will come in time.

The COM port in the script must be configured for the "USB Serial Port" device that represents your RetroTink 4K.

The Retrotink 4K serial commands can be found in the wiki: https://consolemods.org/wiki/AV:RetroTINK-4K#USB_Serial_Configuration

Detected devices will be listed with a (*) beside their name, this relies on the switch itself detecting them which it doesn't always do due mostly the fault of the devices (*cough* PS2 *cough*).

The current active device will be listed with a (A) beside its name.

If input names have been configured in the WebUI (http://gcswitch.local) then these will be retrieved and used as the button labels.

This can be left running on whatever PC is plugged into the Retrotink 4K in "fully automatic" mode if wanted and it'll just silently handle auto-switching of the Profiles as and when the input on the switch changes.
