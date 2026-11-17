"""DEMONEXT Power Distribution Unit interface class

Class to implement remotely operate a power distribution units (PDUs).

Two classes are implemented:
    
RaritanPDU:
    Operates a Raritan PXO PDUs using Raritan's JSON-RPC API python bindings.  
    
    Designed for a Raritan PXO unit with 4 unmetered outlets and 1 metered 
    inlet, which makes this relatively minimalist compared to the range of 
    PDU functions in higher models.  Would require updates if the PDU model 
    has individually-metered outlets or multiple inlets.

PegasusPDU:
    Operates a Pegasus Astro Power Box using its native ASCOM interface
    which is derived from the ASCOM Switch base class. This code is for 
    the Power Box Advanced Gen 2, which has 4x 12VDC outputs that are 
    toggled as a block not individually.  If updating to a Gen 3 PBA unit, 
    will need different switch mapping.

The structure of this module allows us to add other PDU system (e.g.,
those that use generic HTTP or RPC instead of a product-specific
module) at a later date.  The class names include the PDU maker.


Author:
   R. Pogge, OSU Astronomy Dept.
   pogge.1@osu.edu
   2024 Dec 9

Modification History:
    2024 Dec 09 - first version [rwp/osu]
    2024 Dec 10 - added using yaml for reading config, changed __init__() [rwp]
    2024 Dec 11 - choice of cfgFile or kwargs [rwp/osu]
    2024 Dec 15 - many changes to integrate into the demonext module [rwp/osu]
    2024 Dec 17 - added logging system (python logging facility) [rwp/osu]
    2024 Dec 29 - cleaned up docstrings [rwp/osu]
    2024 Dec 31 - bug check and cleanup with spyder [rwp/osu]
    2026 Nov 16 - start of PegasusPDU class development [rwp/osu]
    
"""

import os
import time

# Raritan JSON-RPC API python client bindings

from raritan import rpc
from raritan.rpc import pdumodel
from raritan.rpc import peripheral

# win23com client Dispatch() for ASCOM (Pegasus Power Box)

from win32com.client import Dispatch

# pathlib for path handling

from pathlib import Path

# yaml for configuration file parsing

import yaml

# logging

import logging
logger = logging.getLogger("PDU") 

# RaritanPDU Class

class RaritanPDU:
    """Raritan PDU JSON-RPC interface class"""

    
    def __init__(self,*args,**kwargs):
        """
        Initialize the Raritan JSON-RPC interface and retrieve info we need to query
        and command the Raritan PDU.
        
        Parameters
        ----------
        *args : 
            cfgFile : string
                name of a YAML configuration file (including path) 
            
        **kwargs :
            ipaddr : string
                IP address of the Raritan PDU
            userid : string
                username on the Raritan
            passwd : string
                password for username
            nocert : boolean 
                True for no certificate, False for require certificate
            outlets : string list
                names to assign to outlets 1..N
            timeout : integer
                timeout in seconds for the RPC connection (default: 10s)

        Raises
        ------
        ValueError
            Raised if any arguments are invalid.
        RuntimeError
            Raised if the config file cannot be opened, or if the 
            PDU cannot be accessed.

        Returns
        -------
        None.

        Description
        -----------
        Raritan PDU connection info (IP address, user, and password)
        is expected either in the YAML runtime configuration file as the "pdu" entry
        or given with the kwargs in the code.
      
        The constructor establishes the RPC agent, model, and peripheral manager
        needed to access Raritan PDU functions and data by class member functions and properties.

        If a config file is given as the argument, it expects a YAML formatted file
        with the "pdu" dictionary entry with 3 required parameters:
            ipAddr - the IP address
            userID - the username on the Raritan (should be operator instead of admin)
            passwd - the password for userID
        plus 3 optional parameters:
            disableCert - disable certification verification.  default: True
            outletNames - list of outlet assignments in order from outlet 1..4. default: get from Raritan
            timeout - RPC connection timeout in seconds (default: 10s)

        If no config file is given, kwargs are used to set the parameters, of which
        addr, user, and pass are *required*

        If no config file or valid kwargs are given, we try a default configuration file
        in $HOME/.demonext/config/raritanPDU.txt as a last resort.      

        """

        haveOutNames = False
        self.noCert = True
        self.pduInfo = {}
        self.timeout = 10 # default timeout for rpc
        self.switchDelay = 1 # delay in seconds for outlet switch events
        self.verbose = False # turn on very verbose logging
        
        # Argument options from nothing, a config file, or individual keywords
        
        if len(args) > 0:
            cfgFile = args[0]

        elif len(kwargs) > 0:
            cfgFile = None
            for key, val in kwargs.items():
                if key.lower() == "ipaddr":
                    self.pduAddr = val
                elif key.lower() == "userid":
                    self.pduUser = val
                elif key.lower() == "passwd":
                    self.pduPass = val
                elif key.lower() == "nocert":
                    self.noCert = val
                elif key.lower() == "outlets":
                    self.outletIDs = val
                    haveOutNames = True
                elif key.lower() == "timeout":
                    self.timeout = val
                else:
                    msg = f"Unrecognized kwarg {key}, must be [ipaddr,userid,passwd,nocert,timeout,outlets]"
                    logger.exception(msg)
                    raise ValueError(msg)
        
        else:
            # default config file
            cfgFile = str(Path.home() / ".demonext/config/raritanPDU.txt")

        if cfgFile is not None:
            if os.path.exists(cfgFile):
                with open(cfgFile,"r") as stream:
                    try:
                        config = yaml.safe_load(stream)
                    except yaml.YAMLError as exp:
                        msg = f"Cannot open Raritan runtime configuration file {cfgFile}: {exp}"
                        logger.exception(msg)
                        raise RuntimeError(msg)

                self.pduInfo = config["pdu"]
            
                self.pduAddr = self.pduInfo["ipAddr"]
                self.pduUser = self.pduInfo["userID"]
                self.pduPass = self.pduInfo["passwd"]
                try:
                    self.timeout = self.pduInfo["timeout"]
                except:
                    self.timeout = 10
                    
                try:
                    self.noCert = self.pduInfo["disableCert"]
                except:
                    self.noCert = True

                try:
                    self.outletIDs = self.pduInfo["outletNames"]
                    haveOutNames = True
                except:
                    haveOutNames = False

            else:
                msg = f"Raritan runtime configuration file {cfgFile} does not exist"
                logger.exception(msg)
                raise RuntimeError(msg)
        
        # Instantiate a Raritan RPC agent

        try:
            self.agent = rpc.Agent("http",self.pduAddr,self.pduUser,self.pduPass,
                                   disable_certificate_verification=self.noCert,
                                   timeout=self.timeout)
        except Exception as exp:
            msg = f"Cannot open Raritan PDU RPC agent: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)

        # we have an agent, get PDU model and peripheral device instances

        try:
            self.pdu = pdumodel.Pdu("/model/pdu/0",self.agent)
        except Exception as exp:
            msg = f"Cannot load Raritan pdumodel: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
        
        try:
            self.pdm = peripheral.DeviceManager("/model/peripheraldevicemanager",self.agent)
        except Exception as exp:
            msg = f"Cannot instantiate Raritan PDU peripheral device manager: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
        
        # PDU inlets 

        try:
            self.inlets = self.pdu.getInlets()
        except Exception as exp:
            msg = f"Cannot get Raritan PDU inlets: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)

        self.numInlets = len(self.inlets)
        if self.numInlets > 0:
            try:
                self.inletSensors = self.inlets[0].getSensors()
            except Exception as exp:
                msg = f"Cannot get Raritan PDU inlet sensors: {exp}"
                logger.exception(msg)
                raise RuntimeError(msg)
        else:
            self.inletSensors = None
        
        # PDU outlets

        try:
            self.outlets = self.pdu.getOutlets()
        except Exception as exp:
            msg = f"Cannot get Raritan PDU outlets: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
        
        self.numOutlets = len(self.outlets)
        if self.numOutlets > 0:
            self.outletNames = {}  # outlet name bindings dictionary
            if haveOutNames: # config file has outlet name assignments that override internal PDU names
                for i in range(self.numOutlets):
                    tempID = self.outletIDs[i]
                    if len(tempID) > 0:
                        self.outletNames[tempID.lower()] = i + 1
                    else:
                        self.outletNames[f"Outlet{i+1}"] = i + 1
            else:
                self.outletIDs = []     # use labels assigned to each outlet on the Raritan proper
                
            self.outDelay = []  # outlet power cycle delay in seconds
            self.outState = []  # outlet state sensor objects - query with .getState().value
            iOut = 0
            for outlet in self.outlets:
                iOut += 1
                outSet = outlet.getSettings()

                # no outlet names in the config file, get from the PDU
                
                if not haveOutNames:
                    tempID = outSet.name
                    if len(tempID) == 0:
                        tempID = f"outlet{iOut}"
                    self.outletIDs.append(tempID)
                    self.outletNames[tempID.lower()]=iOut

                # outlet power cycle delay in seconds

                self.outDelay.append(outSet.cycleDelay)

                # outlet on/off sensors

                outSens = outlet.getSensors()
                self.outState.append(outSens.outletState)
                
        else:
            self.outletIDs = None
            self.outDelay = None
            self.outState = None
                
        # PDU peripheral device slots - assume one temp/humidity sensor connected

        try:
            slots = self.pdm.getDeviceSlots()
            self.Temp = slots[0].getDevice() # peripheral temperature sensor device
            self.RH = slots[1].getDevice()   # peripheral humidity sensor device
        except Exception as exp:
            msg = f"Cannot read Raritan PDU peripheral devices: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
        
        # Useful boolean translation dictionaries

        self.OnOff = {True:"On",False:"Off"}
        self.YesNo = {True:"Yes",False:"No"}
        
        


    # Methods

    #------------------------------------
    #
    # low-level methods
    #
    # get/set methods, work at numerical address level
    # See high-level methods to work with named outlets
    #
    
    def getInletData(self):
        """
        Read the Raritan PDU inlent sensors
        
        Returns
        -------
        inletData : dict
            dictionary with the Raritan PDU inlet sensor data
        
        Description
        -----------
        Reads the Raritan PDU inlet sensors and returns inletData dictionary
        with the data.  If a sensor read fails, reports -99.99 (not
        read) for value.

        See also getInletFITS()
        """

        inletData = {}
        try:
            inletData["voltage"] = self.inletSensors.voltage.getReading().value # V
        except:
            inletData["voltage"] = -99.99
        
        try:
            inletData["current"] = self.inletSensors.current.getReading().value # A
        except:
            inletData["current"] = -99.99

        try:
            inletData["peakCurrent"] = self.inletSensors.peakCurrent.getReading().value # A
        except:
            inletData["peakCurrent"] = -99.99

        try:
            inletData["power"] = self.inletSensors.activePower.getReading().value # W
        except:
            inletData["power"] = -99.99

        try:
            inletData["appPower"] = self.inletSensors.apparentPower.getReading().value # VA
        except:
            inletData["appPower"] = -99.99

        try:
            inletData["Energy"] = self.inletSensors.activeEnergy.getReading().value/1000.0 # convert to kWh
        except:
            inletData["Energy"] = -99.99

        try:
            inletData["lineFrequency"] = self.inletSensors.lineFrequency.getReading().value # Hz
        except:
            inletData["lineFrequency"] = -99.99

        return inletData


    def getInletFITS(self):
        """
        Read the Raritan PDU inlet sensors and return a FITS dictionary
        
        Returns
        -------
        inletData : dict
            dictionary of Raritan PDU inlet sensor data with FITS-style keywords.
            
        Description
        -----------
        Reads the Raritan PDU inlet sensors and returns inletFITS dictionary
        with the data in FITS keyword/value pairs ready for inserting
        to an image FITS header.  If a sensor read fails, reports
        -99.99 (not read) for value.
        """

        inletData = {}

        try:
            inletData["PDU_VAC"] = self.inletSensors.voltage.getReading().value
        except:
            inletData["PDU_VAC"] = -99.99
        
        try:
            inletData["PDU_AMPS"] = self.inletSensors.current.getReading().value
        except:
            inletData["PDU_AMPS"] = -99.99

        try:
            inletData["PDU_PEAK"] = self.inletSensors.peakCurrent.getReading().value
        except:
            inletData["PDU_PEAK"] = -99.99

        try:
            inletData["PDU_WATT"] = self.inletSensors.activePower.getReading().value
        except:
            inletData["PDU_WATT"] = -99.99

        try:
            inletData["PDU_VA"] = self.inletSensors.apparentPower.getReading().value
        except:
            inletData["PDU_VA"] = -99.99

        try:
            inletData["PDU_KWH"] = self.inletSensors.activeEnergy.getReading().value/1000.0 # convert to kWh
        except:
            inletData["PDU_KWH"] = -99.99

        try:
            inletData["PDU_FREQ"] = self.inletSensors.lineFrequency.getReading().value
        except:
            inletData["PDU_FREQ"] = -99.99

        return inletData

            
    def readEnv(self):
        """
        Read peripheral environmental sensors connected to the Raritan PDU

        Returns
        -------
        airTemp : float
            air temperature in degrees C.
        airRH : float
            air relative humidity in percent (%).
            
        Description
        -----------
        Reads the temperature and pressure sensors, returns airTemp in
        degrees C and airRH in % relative humidity.  Returns -99.99 if
        a sensor read fails.

        See also readEnvFITS(), readTemp(), readRH()
        """
    
        try:
            airTemp = self.Temp.device.getReading().value
        except:
            airTemp = -99.99

        try:
            airRH = self.RH.device.getReading().value
        except:
            airRH = -99.99

        if self.verbose: logger.debug(f"Read Raritan sensors: T={airTemp:2f}C RH={airRH:.2f}%")
        
        return airTemp, airRH
    

    def getEnvFITS(self):
        """
        Read Raritan PDU peripheral environmental sensors and return FITS cards

        Returns
        -------
        pduEnv : dict
            dictionary with the temperature and relative humidity as FITS header cards.

        Description
        -----------
        Reads the temperature and pressure sensors, returns A FITS
        header dictionary pduEnv with PDU_TEMP and PDU_RH as
        FITS-ready keyword/value pairs.  Returns -99.99 if sensor
        reads fail.

        See also readEnv(), readTemp(), readRH()
        """

        pduEnv = {}
        try:
            pduEnv["PDU_TEMP"] = self.Temp.device.getReading().value
        except:
            pduEnv["PDU_TEMP"] = -99.99

        try:
            pduEnv["PDU_RH"] = self.RH.device.getReading().value
        except:
            pduEnv["PDU_RH"] = -99.99
            
        return pduEnv
        

    def readTemp(self):
        """
        Read the Raritan DU peripheral temperature sensor

        Returns
        -------
        airTemp : float
            air temperature in degrees C.

        Description
        -----------
        Returns a float with the Raritan PDU peripheral temperature sensor
        reading in degrees C or -99.99 if sensor read failed

        See also readEnv(), readRH()
        """
    
        try:
            airTemp = self.Temp.device.getReading().value
        except:
            airTemp = -99.99

        if self.verbose: logger.debug(f"Read Raritan sensors: T={airTemp:2f}C")
        
        return airTemp



    def readRH(self):
        """
        Read the Raritan PDU peripheral relative humidity sensor

        Returns
        -------
        airRH : float
            relative humidity in percent (%).

        Returns a float with the Raritan PDU peripheral relative humidity
        sensor reading in percent or -99.99 if sensor read failed.

        See also readEnv(), readTemp()
        """
        try:
            airRH = self.RH.device.getReading().value
        except:
            airRH = -99.99

        if self.verbose: logger.debug(f"Read Raritan sensors: RH={airRH:2f}%")
        
        return airRH

    
    def getOutlet(self,iOut):
        """
        Get the power status of a Raritan PDU outlet by number

        Parameters
        ----------
        iOut : integer
            Outlet number, 1..numOutlets.

        Raises
        ------
        ValueError
            if an invalid outlet number is given.
        RuntimeError
            if it cannot read the PDU outlet status.

        Returns
        -------
        bool
            Power status of the outlet: True=On, False=Off.

        See also setOutlet() and cycleOutlet(), high-level: isOn(), isOff(), outlets()    
        """

        if iOut < 1 or iOut > self.numOutlets:
            msg = f"Raritan outlet number {iOut} invalid: must be 1..{self.numOutets}"
            logger.exception(msg)
            raise ValueError(msg)

        try:
            status = self.outState[iOut-1].getState().value
            if status:
                return True
            else:
                return False
        except Exception as exp:
            msg = f"Cannot read state of Raritan outlet {iOut}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
        

    def setOutlet(self,iOut,turnOn):
        """
        Set the power state of a Raritan PDU outlet by number

        Parameters
        ----------
        iOut : integer
            Outlet number 1..numOutlets
        turnOn : boolean
            True to turn outlet ON, False to turn outlet OFF.

        Raises
        ------
        ValueError
            raised if an invalid outlet number is given.
        RuntimeError
            raised if it cannot set or get the Raritan PDU outlet state

        Returns
        -------
        bool
            State of the outlet after the requested operation, True=ON, False=OFF.

        Description
        -----------
        Sets the power state of an outlet by number.  The switching operation
        on the Raritan PDU is not blocking, so we wait switchDelay seconds
        before verifying the status after changing the power state.

        See also getOutlet() and cycleOutlet(), high-level: switch(), cycle()
        """

        if iOut < 1 or iOut > self.numOutlets:
            msg = f"Outlet number {iOut} invalid: must be 1..{self.numOutets}"
            logger.exception(msg)
            raise ValueError(msg)

        outlet = self.outlets[iOut-1]

        if turnOn:
            try:
                outlet.setPowerState(outlet.PowerState.PS_ON)
            except Exception as exp:
                msg = f"Cannot switch Raritan outlet {iOut} ON: {exp}"
                logger.exception(msg)
                raise RuntimeError(msg)
        else:
            try:
                outlet.setPowerState(outlet.PowerState.PS_OFF)
            except Exception as exp:
                msg = f"Cannot switch Raritan outlet {iOut} OFF: {exp}"
                logger.exception(msg)
                raise RuntimeError(msg)
                
        time.sleep(self.switchDelay)

        try:
            status = self.outState[iOut-1].getState().value
            if status:
                logger.info(f"Raritan outlet {iOut} switched ON")
                return True
            else:
                logger.info(f"Raritan outlet {iOut} switched OFF")
                return False
            
        except Exception as exp:
            msg = f"Cannot read power state of Raritan outlet {iOut}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)


    def cycleOutlet(self,iOut):
        """
        Power cycle a Raritan PDU outlet by number

        Parameters
        ----------
        iOut : integer
            Outlet number, 1..numOutlets.

        Raises
        ------
        ValueError
            raised if an invalid outlet number given.
        RuntimeError
            raised if it cannot power cycle or query the outlet.

        Returns
        -------
        status : boolean
            Power status of the outlet after cycling, True=ON, False=OFF.
            
        Description
        -----------
        Power cycles the outlet on the Raritan PDU.  Internally the Raritan PDU has a preset
        delay between powering off then back on of 10 seconds, but this
        can be changed on the Raritan PDU admin interface.  We read this as the
        outDelay property read from the PRaritan DU when the class was
        instantiated.  
        
        This function will pause for a time of outDelay plus twice
        switchDelay (once each for off then on) to wait give the Raritan PDU
        time to complete the power cycling operation before reading
        the outlet sensor to verify the post-cycling power state. For
        typical default values, power cycling an outlet takes about 12 seconds.
        
        See also: cycle()
        
        """
        
        if iOut < 1 or iOut > self.numOutlets:
            msg = f"Raritan outlet number {iOut} invalid: must be 1..{self.numOutets}"
            logger.exception(msg)
            raise ValueError(msg)

        outlet = self.outlets[iOut-1]

        try:
            outlet.cyclePowerState()
            time.sleep(self.outDelay[iOut-1]+2*self.switchDelay) # wait cycle delay + 2 switch delays (on/off)
            status = self.getOutlet(iOut)
            logger.info(f"Raritan outlet {iOut} power cycled, state={self.OnOff[status]}")
            return status

        except Exception as exp:
            msg = f"Cannot power cycle Raritan outlet {iOut}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)

        
    #------------------------------------
    # 
    # High-level methods
    #
    
    def outlets(self):
        """
        Return the power state of all PDU outlets as a list of booleans

        Raises
        ------
        RuntimeError
            raised if it cannot query the outlet states on the Raritan PDU.

        Returns
        -------
        states : boolean list
            List of booleans with the current power state of each outlet, 
            True=ON, False=OFF.
            
        Example
        -------
        >>> config.outlets()
        >>> [T,T,F,F]
        
        indicates that outlets 1 and 2 are on, 3 and 4 are off.

        See also: isOn(), isOff(), printOutlet(), low-level: getOutlet()    
        """

        states = []
        for iOut in range(self.numOutlets):
            try:
                status = self.outState[iOut-1].getState().value
                if status:
                    states.append(True)
                else:
                    states.append(False)
            except Exception as exp:
                msg = f"Cannot read power state of Raritan outlet {iOut+1}: {exp}"
                logger.exception(msg)
                states = []
                raise RuntimeError(msg)
        
        return states

    
    def switch(self,outName,action):
        """
        Switch named Raritan outlet on or off.

        Parameters
        ----------
        outName : string
            name of an outlet.
        action : string
            "on" to switch the outlet on, "off" to switch it off.
            action is case-insensitive.

        Raises
        ------
        ValueError
            Raised if an outlet name or action string is invald.
        RuntimeError
            Raised if it cannot complete or confirm the requested operation.

        Returns
        -------
        status : boolean
            Power state of the outlet after the action, True=ON, False=OFF.
            
        Description
        -----------
        Uses the outletNames dictionary of assigned outlet names built on
        class initialization to as the list of allowed outlet names.  All
        outlet names are forced to lowercase on initialization so the
        outName is case-insensitive.
        
        We use the setOutlet() method defined in this class for to execute
        the requested switching operation.
        
        See also: cycle(), isOn(), isOff(), low-level: setOutlet(), cycleOutlet()
        """

        outID = outName.lower()
        if outID not in self.outletNames:
            msg = f"switch(): Raritan outlet {outName} invalid, must be one of {self.outletNames}"
            logger.exception(msg)
            raise ValueError(msg)
        iOut = self.outletNames[outID]

        act = action.lower()
        if act not in ['on','off']:
            msg = f"switch(): Raritan action {action} invalid: must be on or off"
            logger.exception(msg)
            raise ValueError(msg)
         
        try:
            status = self.setOutlet(iOut,act=='on')
            logger.info(f"Raritan outlet {outName} switched {self.OnOff[status]}")
            return status
                        
        except Exception as exp:
            msg = f"Cannot switch Raritan outlet {outName} {action.upper()}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)

    
    def cycle(self,outName):
        """
        Power cycle an Raritan outlet by name

        Parameters
        ----------
        outName : string
            valid outlet name (in outletNames property).

        Raises
        ------
        ValueError
            Raised if an invalid outlet name is given.
        RuntimeError
            Raised if it cannot complete or confirm power cycling.

        Returns
        -------
        status : boolean
            Power state of the outlet after power cycling, True=ON, False=OFF.

        Description
        -----------
        Uses the outletNames dictionary of assigned outlet names built on
        class initialization to as the list of allowed outlet names.  All
        outlet names are forced to lowercase on initialization so the
        outName is case-insensitive.
        
        The PDU has an internal delay of usually 10s (but user-definable on
        by the PDU web admin interface) between power off then on cycle, which
        we query and store in the outDelay property.  We wait for outDelay
        plus twice the switchDelay time (once each for off and on) to make
        sure the operation completes before querying the outlet state to confirm
        the power state after power cylcing. For typical default settings
        this amounts to a power cycling time of about 12 seconds.
        
        See also: switch(), isOn(), isOff(), low-level: cycleOutlet(), setOutlet()
        """
        outID = outName.lower()
        if outID not in self.outletNames:
            msg = f"cycle(): Raritan outlet {outName} invalid, must be one of {self.outletNames}"
            logger.exception(msg)
            raise ValueError(msg)
        iOut = self.outletNames[outID]

        try:
            self.outlets[iOut-1].cyclePowerState()
            time.sleep(self.outDelay[iOut-1]+2*self.switchDelay) # wait cycle delay + 2 switch delays (on/off)
            status = self.getOutlet(iOut)
            logger.info(f"Raritan outlet {outName} outlet power cycled, state={self.OnOff[status]}")
            return status
        except Exception as exp:
            msg = f"Cannot power cycle Raritan outlet {outName}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)

        
    def isOn(self,outName):
        """
        Is the named outlet switched on?

        Parameters
        ----------
        outName : string
            valid outlet name (in the outletNames property).

        Raises
        ------
        ValueError
            Raised if an invalid outlet name is given.
        RuntimeError
            Raised if the outlet state cannot be read.

        Returns
        -------
        boolean
            True if the outlet is ON, False if the outlet is OFF.

        Description
        -----------
        Uses the outletNames dictionary of assigned outlet names built on
        class initialization to as the list of allowed outlet names.  All
        outlet names are forced to lowercase on initialization so the
        outName is case-insensitive.
    
        Returns boolean True if the outlet is On, False if it is Off.

        See also isOff() for the reverse test, outlets(), low-level: getOutlet()
        """

        outID = outName.lower()
        if outID not in self.outletNames:
            msg = f"isOn(): Raritan {outName} invalid, must be one of {self.outletNames}"
            logger.exception(msg)
            raise ValueError(msg)
        try:
            return self.getOutlet(self.outletNames[outID])
        except Exception as exp:
            msg = f"Cannot read Raritan outlet {outName} state: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)


    def isOff(self,outName):
        """
        Is the named outlet switched off?

        Parameters
        ----------
        outName : string
            valid outlet name (in the outletNames property).

        Raises
        ------
        ValueError
            Raised if an invalid outlet name is given.
        RuntimeError
            Raised if the outlet state cannot be read.

        Returns
        -------
        boolean
            True if the outlet is OFF, False if the outlet is ON.

        Opposite of isOn() method for convenience.

        See also isOn() for the reverse test, outlets(), low-level: getOutlet()
        """

        outID = outName.lower()
        if outID not in self.outletNames:
            msg = f"isOff(): Raritan {outName} invalid, must be one of {self.outletNames}"
            logger.exception(msg)
            raise ValueError(msg)
        try:
            return not self.getOutlet(self.outletNames[outID])
        except Exception as exp:
            msg = f"Cannot read Raritan outlet {outName} state: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)


    #------------------------------------
    #
    # Formatted printing functions
    #
    # these are all non-logging
    
    def printOutlets(self):
        """
        Print formatted status of all Raritan PDU outlets

        Returns
        -------
        None.

        Prints a human-readable list of the current output status.

        See also outlets(), low-level: getOutlet()
        """
    
        print("\nRaritan PDU outlet status:")
        for i in range(self.numOutlets):
            outID = self.outletIDs[i]
            delay = self.outDelay[i]
            outlet = self.outlets[i]
            try:
                outOnOff = self.OnOff[self.outState[i].getState().value]
            except:
                outOnOff = "??"

            if outlet.getSettings().startupState == pdumodel.Outlet.StartupState.SS_ON:
                defOnOff = "ON"
            else:
                defOnOff = "OFF"
            print(f"  Outlet {i+1}: {outID}")
            print(f"          Power: {outOnOff}")
            print(f"        Startup: {defOnOff}")
            print(f"    Cycle Delay: {delay} seconds\n")


    def printInlet(self):
        """
        Print formatted status of the Raritan PDU inlet power sensors

        Raises
        ------
        RuntimeError
            Raised if it cannot read the Raritan PDU inlet sensors.

        Returns
        -------
        None.

        Prints a human-readable summary of relevant Raritan PDU inlet sensor data

        See also low-level getInlet(), getInletFITS()  
        """

        try:
            rmsV = self.inletSensors.voltage.getReading().value
            freq = self.inletSensors.lineFrequency.getReading().value
            rmsI = self.inletSensors.current.getReading().value
            peakI = self.inletSensors.peakCurrent.getReading().value
            power = self.inletSensors.current.getReading().value
            VA = self.inletSensors.apparentPower.getReading().value
            kWh = self.inletSensors.activeEnergy.getReading().value/1000.0 # conver to kWh
        except Exception as exp:
            raise RuntimeError(f"Cannot read PDU inlet sensor data: {exp}")
            
        print("\nRaritan PDU Inlet Sensors:")
        print(f"    RMS Voltage: {rmsV:.2f} VAC, {freq:.2f} Hz")
        print(f"    RMS Current: {rmsI:.3f} A, {peakI:.3f} A peak")
        print(f"          Power: {power:.2f} W")
        print(f"  Active Energy: {kWh:.2f} kWh")
        print(f" Apparent Power: {VA:.1f} VA")

    
    def printEnv(self): 
        """
        Print PDU peripheral sensor data

        Raises
        ------
        RuntimeError
            raised if it cannot read the Raritan PDU peripheral sensors.

        Returns
        -------
        None.

        Prints a human-readable readout summary of the Raritan PDU periperal
        temperature and humidity sensors.

        See also readEnv(), readTemp(), readRH(), getEnvFITS()
        """

        try:
            temp,rh = self.readEnv()
        except Exception as exp:
            raise RuntimeError(f"Cannot read Raritan PDU environmental sensor data: {exp}")
        
        print("\nRaritan PDU peripheral sensors:")
        print(f"  Temperature: {temp:.2f}C")
        print(f"     Humidity: {rh:.2f}%")

        
    def printMetaData(self):
        """
        Print the Raritan PDU inlet metdata

        Raises
        ------
        RuntimeError
            Raised if it cannot read the Raritan PDU inlet metadata.

        Returns
        -------
        None.

        Prints a human-readable summary of the Raritan PDU metadata.  The metadata
        is presented in JSON format.
        """

        try:
            md = self.inlets[0].getMetaData()
        except Exception as exp:
            raise RuntimeError(f"Cannot retrieve Raritan PDU metadata: {exp}")

        print("\nRaritan PDU inlet metadata:")
        print(md)

###################################################################################

# PegasusPDU Class

class PegasusPDU:
    """Pegasus Astro Power Box PDU ASCOM remote interface class"""

    
    def __init__(self,*args,**kwargs):
        """
        Initialize the Pegasus Power Box ASCOM interface and retrieve info we need to query
        and command the Pegasus PDU.
        
        Parameters
        ----------
        *args : 
            cfgFile : string
                name of a YAML configuration file (including path) 
            
        **kwargs :
            ascomID : string
                ASCOM device ID of the Pegasus Power Box unit
            switches: dict
                Dictionary of switches - {addr:"ID"}
            sensors: dict
                Dictionary of sensors - {addr:"ID"}
            switchDelay: float
                time delay following switching operations in seconds

        Raises
        ------
        ValueError
            Raised if any arguments are invalid.
        RuntimeError
            Raised if the config file cannot be opened, or if the 
            Power Box unit cannot be accessed.

        Returns
        -------
        None.

        Description
        -----------
        Pegasus Power Box connection info (ASCOM ID) is expected either in the 
        YAML runtime configuration file or given with the kwargs in the code.
      
        The constructor establishes the ASCOM object and parameters needed
        to operate the Pegasus Power Bo PDU functions and data by class member
        functions and properties.

        If a config file is given as the argument, it expects a YAML formatted file
        with the "ppb" dictionary entry :
            ascomID - the ASCOM ID of the Pegasus Power Box
            switchDelay - time delay following switching operation in seconds
            switches - dictionary of switches, format {addr:"switchID"}
            sensors - dictionary of sensors, format {addr:"sensorID"}

        If no config file is given, kwargs are used to set the parameters

        If no config file or valid kwargs are given, we try a default configuration file
        in $HOME/.demonext/config/pegasusPDU.txt as a last resort.      

        """

        self.ppbInfo = {}
        self.switchDelay = 1 # delay in seconds for outlet switch events
        self.verbose = False # turn on very verbose logging
        
        # Default switches and sensors for the Pegasus Power Box Advance Gen 2
        
        self.switches = {"Quad":0,"Adj":1,"DewA":2,"DewB":3,"AutoDew":7,"USB2":8}
        self.sensors = {"VDC":4,"Amps":5,"Watts":6}
        
        self.Name = ""
        self.maxAddr = 0
        self.numAddr = 0
        
        # Argument options from nothing, a config file, or individual keywords
        
        if len(args) > 0:
            cfgFile = args[0]

        elif len(kwargs) > 0:
            cfgFile = None
            for key, val in kwargs.items():
                if key.lower() == "ascomid":
                    self.ascomID = val
                
                elif key.lower() == "switchDelay":
                    try:
                        self.switchDelay = float(val)
                    except Exception as exp:
                        msg = f"Invalid switchDelay, must be float - {exp}"
                        logger.exception(msg)
                        raise ValueError(msg)
                
                elif key.lower() == "switches":
                    self.switches = val
                    
                elif key.lower() == "sensors":
                    self.sensors = val

                else:
                    msg = f"Unrecognized kwarg {key}, must be [ipaddr,userid,passwd,nocert,timeout,outlets]"
                    logger.exception(msg)
                    raise ValueError(msg)
        
        else:
            # default config file
            cfgFile = str(Path.home() / ".demonext/config/pegasusPDU.txt")

        if cfgFile is not None:
            if os.path.exists(cfgFile):
                with open(cfgFile,"r") as stream:
                    try:
                        config = yaml.safe_load(stream)
                    except yaml.YAMLError as exp:
                        msg = f"Cannot open Raritan runtime configuration file {cfgFile}: {exp}"
                        logger.exception(msg)
                        raise RuntimeError(msg)

                self.ppbInfo = config["ppb"]
            
                self.ascomID = self.ppbInfo["ipAddr"]
                    
                try:
                    self.switches = self.pdbInfo["switches"]
                except:
                    self.switches = {"Quad":0,"Adj":1,"DewA":2,"DewB":3,"AutoDew":7,"USB2":8}
                    
                try:
                    self.sensors = self.pdbInfo["sensors"]
                except:
                    self.sensors = {"VDC":4,"Amps":5,"Watts":6}
                    
                try:
                    self.switchDelay = self.pduInfo["switchDelay"]
                except:
                    self.switchDelay = 1.0

            else:
                msg = f"Pegasus runtime configuration file {cfgFile} does not exist"
                logger.exception(msg)
                raise RuntimeError(msg)
        
        # Instantiate a win32 com dispatcher

        try:
            self.ppb = Dispatch(self.ascomID)
        except Exception as exp:
            msg = f"Cannot init Pegasus Power Box ASCOM object {self.ascomID}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)

        # Connect the Pegasus unit
        
        try:
            self.ppb.Connected = True
            self.Name = self.ppb.Name
            self.numAddr = self.ppb.MaxSwitch
            self.maxAddr = self.numAddr - 1
            logger.info(f"Connected to {self.Name}, {self.numAddr} switches/sensors")
            
        except Exception as exp:
            msg = f"Cannot connect to Pegasus Power Box: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
            
        # Useful boolean translation dictionaries

        self.OnOff = {True:"On",False:"Off"}
        self.YesNo = {True:"Yes",False:"No"}
        

    # Methods

    def getInletData(self):
        '''
        Retrieve Pegasus PDU power input sensor data (VDC, amps, watts)

        Returns
        -------
        inletData : dict
            dictionary with the Pegasus Power Box inlet data (VDC, amps, watts)
            
        Description
        -----------
        Reads the values of the Pegasus PDU input sensors (volts, current, 
        power) and returns a dictionary with the current values. -99.99 is 
        the no-read value.
        
        See also getInletFITS()
        
        '''
        
        inletData = {}
        
        for sensID,addr in self.sensors.items():
            try:
                inletData[sensID] = self.ppb.GetSwitchValue(addr)
            except:
                inletData[sensID] = -99.99
        
        return inletData
    

    def getInletFITS(self):
        '''
        Retrieve power input sensor data and return a FITS keyword dictionary 
    
        Returns
        -------
        inletFITS : dict
            dictionary of Pegasus Power Obx inlet data as FITS-style keywords
        
        Description
        -----------
        Reads the values of the Pegasus PDU input sensors (VDC, amps, watts)
        and retunrs them as FITS keyword=value pairs ready for inclusion in
        FITS headers.  -99.99 is a no-read value.
        
        See also getInletData()
        
        '''
        
        inletFITS = {}
        for sensID,addr in self.sensors.items():
            try:
                inletFITS[f"PB_{sensID}"] = self.ppb.GetSwitchValue(addr)
            except:
                inletFITS[f"PB_{sensID}"] = -99.99
                
        return inletFITS


    def getSwitch(self,swAddr):
        '''
        Get switch state by switch number (0..N)
        
        Parameters
        ----------
        swAddr : int
            Address of the switch.
            
        Returns
        -------
        boolean : 
            Power status of the switch, True=On, False=OFF
        
        See also setSwitch()
        '''
        
        if swAddr < 0 or swAddr > self.maxAddr:
            msg = f"Pegasus Power Box switch {swAddr} invalid, must be 0..{self.maxAddr}"
            logger.exception(msg)
            raise ValueError(msg)
            
        try:
            return self.ppb.getSwitch(swAddr)
        
        except Exception as exp:
            msg = f"Cannot read state of Pegasus Power Box switch {swAddr}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
            
    
    def setSwitch(self,swAddr,turnOn):
        '''
        Set the power state of a Pegasus Power Box switch by number

        Parameters
        ----------
        swAddr : int
            Address of the switch
        turnOn : bool
            True to turn switch ON, False to turn switch OFF

        Returns
        -------
        boolean :
            state of the switch after the requested operation
        
        Description
        -----------
        Sets a switch ON or OFF by number. The operation of the Pegasus 
        Power Box switches is non-blocking, so we wait switchDelay before
        verifying the switch status and reporting the result.
        
        See also getSwitch()
        
        '''
        
        if swAddr < 0 or swAddr > self.maxAddr:
            msg = f"Pegasus Power Box switch {swAddr} invalid, must be 0..{self.maxAddr}"
            logger.exception(msg)
            raise ValueError(msg)

        try:
            self.ppb.SetSwitch(swAddr,turnOn)
            time.sleep(self.switchDelay)
            return self.ppb.GetSwitch(swAddr)
        
        except Exception as exp:
            msg = f"Cannot set Pegasus Power Box switch {swAddr} {'ON' if turnOn else 'OFF'}: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
            
        
    def getSwitches(self):
        '''
        Retrieve the status of all switches on the Pegasus Power Box

        Returns
        -------
        swState : dict
            dictionary of switch status as booleans (True=ON, False=OFF)
            
        Description
        -----------
        Reads all switches and returns their status ON (True) or OFF (False)
        as a dictionary
        
        See also getSwitch() to return state of one switch
        '''
        
        swState = {}
        for swName,swAddr in self.switches.items():
            try:
                swState[swName] = self.ppb.getSwitch(swAddr)
            except Exception as exp:
                msg = f"Cannot read Pegasus Power Box switch {swName} ({swAddr}) state: {exp}"
                logger.exception(msg)
                raise RuntimeError(msg)
                
        return swState
    
    
    def switch(self,swName,action):
        '''
        Turn a named Pegasus Power Box switch on or off
        
        Parameters
        ----------
        swName : string
            Name of the switch (case-insensitive).
        action : string
            switch action, ON or OFF (case-insensitive).

        Returns
        -------
        swState : bool
            Power state of the named switch, True=ON, False=OFF
            
        See also: setSwitch(), getSwitch()
        '''
    
        # validate the requested switch name
        
        if swName.lower() in [x.lower() for x in [*self.switches.keys()]]:
            swAddr = {k.lower():v for k,v in self.switches.items()}[swName.lower()]
        else:
            msg = f"ERROR: Requested Pegasus Power Box switch {swName} is invalid"
            logger.error(msg)
            raise ValueError(msg)
            
        # got a valid switch, perform action if ON or OFF

        if action.upper in ["ON","OFF"]:    
            self.ppb.SetSwitch(swAddr,action.upper()=="ON")
            time.sleep(self.switchDelay)
            return self.ppb.getSwitch(swAddr)
        else:
            msg = f"ERROR: Requested Pegasus Power Box switch action {action} invalid, must be ON or OFF"
            logger.error(msg)
            raise ValueError(msg)
            

    def isOn(self,swName):
        '''
        Is named Pegasus switch on?

        Parameters
        ----------
        swName : string
            Name of the switch to test.

        Returns
        -------
            bool : True if On, False if Off
            
        Description
        -----------
        Tests to see if a named switch is ON.
        
        See also isOff()
        
        '''    
        # validate the requested switch name
        
        if swName.lower() in [x.lower() for x in [*self.switches.keys()]]:
            swAddr = {k.lower():v for k,v in self.switches.items()}[swName.lower()]
        else:
            msg = f"ERROR: Requested Pegasus Power Box switch {swName} is invalid"
            logger.error(msg)
            raise ValueError(msg)
        
        # Test is switch is on
        
        try:
            return self.ppb.GetSwitch(swAddr)
        except Exception as exp:
            msg = f"Cannot read Pegasus Power Box switch {swName} (addr: {swAddr} state: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
            
        
    def isOff(self,swName):
        '''
        Is named Pegasus switch Off?

        Parameters
        ----------
        swName : string
            Name of the switch to test.

        Returns
        -------
            bool : True if Off, False if On
            
        Description
        -----------
        Tests to see if a named switch is Off.
    
        See also isOn()
        
        '''    
        # validate the requested switch name
        
        if swName.lower() in [x.lower() for x in [*self.switches.keys()]]:
            swAddr = {k.lower():v for k,v in self.switches.items()}[swName.lower()]
        else:
            msg = f"ERROR: Requested Pegasus Power Box switch {swName} is invalid"
            logger.error(msg)
            raise ValueError(msg)
        
        # Test if switch is off
        
        try:
            return not self.ppb.GetSwitch(swAddr)
        except Exception as exp:
            msg = f"Cannot read Pegasus Power Box switch {swName} (addr: {swAddr} state: {exp}"
            logger.exception(msg)
            raise RuntimeError(msg)
            

    def printStatus(self):
        '''
        Print a formatted table of the Pegasus Power Box status.

        Returns
        -------
        None
        '''
        
        print("\nPegasus Power Box Status:")
        print(f"  Name: {self.Name}")
        print(f"  Switches/Sensors: {self.numAddr}")
        print("Switches:")
        try:
            swState = self.getSwitches()
            for swName,swAddr in self.switches.items():
                print(f"  Switch {swAddr}: {swName}: {'ON' if swState[swName] else 'OFF'}")
        except Exception as exp:
            print(f"  Cannot read Pegasus switch states: {exp}")
        
        print("Input: ")
        try:
            inData = self.getInletData()
            inStr = "  Power In: "
            for inName in self.sensors:
                inStr += f"{inData[inName]} {inName} "
            print(inStr)
        except Exception as exp:
            print(f"  Cannot read input power data: {exp}")
            
