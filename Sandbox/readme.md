# DEMONEXT 2025 Development Sandbox

Scripts and Jupyter notebooks used for developing the DEMONEXT 2025 system code.

## Development Jupyter Notebooks

### `DEMONEXT Sandbox.ipynb`

Sandbox that I've used to run functional tests of the system while bringing it up on the new Windows 10
computer with the updated software, and to start development of some of the code we'll need. 

Uses the yaml `demonext.txt` configuration file, in `Config/`

### `ImageAnalysis_GuiderClass.ipynb`

Sandbox for testing the `Guider` class in the `demonext` module in `Scripts/`.  A work in progress as of
2025 Jan 25.  Needs images from the Winer DEMONEXT system, see Rick for how to access the repository on
the ASCTech shared DEMONEXT store.

### `Raritan PDU sandbox.ipynb`

Sandbox for learning how to operate the Raritan PXO power distribution unit (PDU) with the 
Raritan JSON-RPC interface.

Uses the yaml `demonext.txt` configuration file, in `Config`

### Development and testing Jupyter notebooks

 * `DEMONEXT StartUp_Shutdown.ipynb` used to develop and document the startup and shutdown procedures.  We use this to startup and shutdown the system during unit-level code testing.
 * `TelescopeSandbox.ipynb` to develop, test, and document telescope operation (`Telescope` class)
 * `CameraSandbox.ipynb` to develop, test, and document camera (`Camera` class) and focuser (`Focuser` class) operations. Note the filter wheel is controlled through the Camera class.  Tests include telescope interaction.
 * `SRO_Sandbox.ipynb` to develop, test, and document code for reading the SRO site weather data and building roll-off roof status (open or closed).
 * `Site_Sandbox.ipynb` to develop and test the `Site` class

The current versions of all these notebooks were live testing during DEMONEXT installation and post-installation verification and alignment at SRO during the week of 2026 March 15-19.

### Demo programs

 * `dxDemo.py` to debug the runtime config file handling and PDU control (`Config` and `RaritanPDU` classes).  Elements of this are now in the growing suite of sandbox notebooks above.
 * `telDemo.py` for live testing of the Telescope class in the lab - superceded by the `TelescopeSandbox.ipynb` notebook

### ToDo

 * port `guider.py` to a new class to implement SV's original "science guiding" mode from the 2016 system in python 3.
 * incorporate an electronic focuser for the guide telescope for later deployment


### raritanPDU

Demonstration class to implement remote control of the Raritan PXO-2402R-A16 Power Distribution
Unit (PDU) that we use to control and monitor AC power for the PC, telescope drive, and instrument
package on DEMONEXT. 

The PDU we use has 4 switchable but not individually metered AC outlets, one metered input, and a 
DX2-T1H1 peripheral RH/Temp sensor module connected to monitor air temperature and humidity
inside the DEMONEXT electronics box through the PDU.

The code communicates with the PDU using the Raritan Xerxus JSON-RPC API python client bindings.  This
requires the `raritan` python module (https://pypi.org/project/raritan/) installed using `pip install raritan`.

#### Important Note

This version of the Raritan PDU code was for initial test and evaluation to see if the PDU would work for
us. For the DEMONEXT flight code use the `RaritanPDU` class implementation in the `config.py` submodule of the 
`demonext` module.  This class will be moved elsewhere on the GitHub repository after we consolidate and release
the flight code.

