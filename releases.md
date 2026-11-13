# pyDEMONEXT Release Notes

**Last Release: 2026 Nov 12**

## VErsion 1.1.0 - 2026 Nov 12
Start of updates for the impending swapout of the FLI CCD camera for a new ZWO sCMOS sensor, and addition of a Pegasus Power Box focal plane instrument PDU and USB hub to replace the power and USB distribution functions of the FLI CCD.
 * `Config/demonext.txt` - adding hooks for new equipment
 * `Config/raritanPDU.txt` - changes for the instrument swap
 * `Config/pegasusPDU.txt` - prototype config file for the Pegasus PDU/Hub
 * `Sandbox/Pegasus_Sandbox.ipynb` - ASCOM interface sandbox for the Pegasus box

ToDo:
 * EAF focuser for the guide telescope (interface through MaxIM)
 * Changes to code to replace "CCD" designation with "Science" for the camera.
 * Lots more to come!

## Version 1.0.3 - 2026 Jun 06
Updates from live testing:
 * `Config/demonext.txt` - added `skyflats` dictionary with parameters needed to automate acqusition of sky flats (work in progress)
 * `demonext/site.py` - added `SUNALT` keyword to FITS header info returned by the `siteTelemetry()` method. Gives sun altitude at start of exposure, useful for sky flats.


## Version 1.0.2 - 2026 May 19
Updates and bug fixes from live testing at SRO:
 * `demonext/camera.py` - fixed bug in `Camera.cooldown()` method that prevented changing the set point.


## Version 1.0.1 - 2026 May 16

Migrated working code from https://github.com/kyleecpi/DEMONEXT after deployment of the telescope and instrumentation at Sierra Remote Observatories in March 2026.

New development will continue in this repository.  This separates the core observatory telescope and instrument control code from the overall director and
scheduler code development led by Kylee Carden at JHU.
