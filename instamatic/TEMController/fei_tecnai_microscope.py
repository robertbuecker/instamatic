import comtypes.client
import atexit
import time
import random
import threading
from numpy import pi

import logging
logger = logging.getLogger(__name__)

from instamatic import config

"""
speed table (deg/s):
1.00: 21.14
0.90: 19.61
0.80: 18.34
0.70: 16.90
0.60: 14.85
0.50: 12.69
0.40: 10.62
0.30: 8.20
0.20: 5.66
0.10: 2.91
0.05: 1.48
0.04: 1.18
0.03: 0.888
0.02: 0.593
0.01: 0.297
"""

# ranges for calib, not actual modes. So it's 'mag' for all mags!
FUNCTION_MODES = {1:'mag', 2:'mag', 3:'mag', 4:'mag', 5:'LAD', 6:'D'}
PROBE_MODES = ['micro', 'nano']
STEM_FOCUS_STRATEGY = ['Intensity', 'Objective', 'StepSize', 'Both']

MagnificationMapping = {
    1: 21,
    2: 28,
    3: 38,
    4: 56,
    5: 75,
    6: 97,
    7: 120,
    8: 170,
    9: 220,
    10 :330,
    11 :420,
    12 :550,
    13 :800,
    14 :1100,
    15 :1500,
    16 :2100,
    17 :1700,
    18 :2500,
    19 :3500,
    20 :5000,
    21 :6500,
    22 :7800,
    23 :9600,
    24 :11500,
    25 :14500,
    26 :19000,
    27 :25000,
    28 :29000,
    29 :50000,
    30 :62000,
    31 :80000,
    32 :100000,
    33 :150000,
    34 :200000,
    35 :240000,
    36 :280000,
    37 :390000,
    38 :490000,
    39 :700000}

CameraLengthMapping = {
    1:  52,
    2:  70,
    3:  100,
    4:  120,
    5:  150,
    6:  200,
    7:  285,
    8:  320,
    9:  520,
    10: 730,
    11: 1000,
    12: 1200,
    13: 1500,
    14: 2000,
    15: 3000,
    16: 6000
    }

CameraLengthMappingLAD = {}

class FEITecnaiMicroscope(object):
    """docstring for FEI microscope"""
    def __init__(self, name = "fei_tecnai_f20_v3"):
        super(FEITecnaiMicroscope, self).__init__()
        
        try:
            comtypes.CoInitializeEx(comtypes.COINIT_MULTITHREADED)
        except WindowsError:
            comtypes.CoInitialize()
            
        print("Philips Tecnai F20 initializing...")
        ## tem interfaces the GUN, stage obj etc but does not communicate with the Instrument objects
        self.tem = comtypes.client.CreateObject("TEMScripting.Instrument.1", comtypes.CLSCTX_ALL)
        ## tecnai does similar things as tem; the difference is not clear for now
        self.tecnai = comtypes.client.CreateObject("Tecnai.Instrument", comtypes.CLSCTX_ALL)
        ## tom interfaces the Instrument, Projection objects
        self.tom = comtypes.client.CreateObject("TEM.Instrument.1", comtypes.CLSCTX_ALL)
        
        ### TEM Status constants
        self.tem_constant = comtypes.client.Constants(self.tem)
        
        self.stage = self.tem.Stage
        self.proj = self.tem.Projection
        self.proj_tom = self.tom.Projection
        self.stage_tom = self.tom.Stage
        self.illu = self.tem.Illumination
        self.illu_tom = self.tom.Illumination
        self.gun = self.tem.GUN
        self.acq = self.tem.Acquisition
        self.stem_tom = self.tom.STEM
        
        t = 0
        while True:
            ht = self.gun.HTValue
            if ht > 0:
                break
            time.sleep(1)
            t += 1
            if t > 3:
                print("Waiting for microscope, t = {}s".format(t))
            if t > 30:
                raise RuntimeError("Cannot establish microscope connection (timeout).")

        logger.info("Microscope connection established")
        atexit.register(self.releaseConnection)

        self.name = name
        self.FUNCTION_MODES = FUNCTION_MODES

        self.FunctionMode_value = 0

        for mode in self.FUNCTION_MODES.values():
            attrname = "range_{}".format(mode)
            try:
                rng = getattr(config.microscope, attrname)
            except AttributeError:
                print("Warning: No magnfication ranges were found for mode `{}` in the config file".format(mode))
            else:
                setattr(self, attrname, rng)
        
        self.goniostopped = self.stage.Status
        
        input("Please select the type of sample stage before moving on.\nPress <ENTER> to continue...")

        # self.Magnification_value = random.choice(self.MAGNIFICATIONS)
        #self.Magnification_value = 2500
        #self.Magnification_value_diff = 300

    def getHTValue(self):
        return self.gun.HTValue
    
    def setHTValue(self, htvalue):
        self.gun.HTValue = htvalue
        
    def getMagnification(self):
        if self.proj_tom.Mode != 1:
            return self.proj.Magnification
        else:
            return self.proj.CameraLength * 1000
    
    def setMagnification(self, value):
        current_mode = self.getFunctionMode()
        
        if current_mode == "D":
            if value not in self.range_D:
                raise IndexError("No such camera length: {}".format(value))
            self.proj.CameraLengthIndex = self.range_D.index(value) + 1
        elif current_mode == "mag":
            if value not in self.range_mag:
                raise IndexError("No such magnification: {}".format(value))
            self.proj.CameraLengthIndex = self.range_ag.index(value) + 1
        elif current_mode == 'LAD':
            raise NotImplementedError('LAD mode currently not supported')
        
    # TODO: stage speed should maybe be handled more cleverly...
    def setStageSpeed(self, value):
        """Value be 0 - 1"""
        if value > 1 or value < 0:
            raise ValueError("setStageSpeed value must be between 0 and 1. Input: {}".format(value))

        self.stage_tom.Speed = value
        
    def getStageSpeed(self):
        return self.stage_tom.Speed
        
    def getStagePosition(self):
        """return numbers in microns, angles in degs."""
        return self.stage.Position.X * 1e6, self.stage.Position.Y * 1e6, self.stage.Position.Z * 1e6, self.stage.Position.A / pi * 180, self.stage.Position.B / pi * 180
    
    def setStagePosition(self, x=None, y=None, z=None, a=None, b=None, wait = True, speed = 1):
        """x, y, z in the system are in unit of meters, angles in radians.
        On a Tecnai (v3), this is a total mess."""
        pos = self.stage.Position
        axis = 0
        
        if speed > 1 or speed < 0:
            raise ValueError("setStageSpeed value must be between 0 and 1. Input: {}".format(speed))
        
        if x is not None:
            pos.X = x * 1e-6
            axis += 1
        if y is not None:
            pos.Y = y * 1e-6
            axis += 2
        if z is not None:
            pos.Z = z * 1e-6
            axis += 4
        if a is not None:
            pos.A = a / 180 * pi
            axis += 8
        if b is not None:
            pos.B = b / 180 * pi
            axis += 16
            
        if speed == 1:
            self.stage.Goto(pos, axis)
        else:
            sp0 = self.stage_tom.Speed
            self.stage_tom.Speed = speed
            if x is not None:
                self.stage_tom.GotoWithSpeed(0, pos.X)
            if y is not None:
                self.stage_tom.GotoWithSpeed(1, pos.Y)
            if z is not None:
                self.stage_tom.GotoWithSpeed(2, pos.Z)
            if a is not None:
                self.stage_tom.GotoWithSpeed(3, pos.A)
            if b is not None:
                raise NotImplementedError('Beta tilt cannot be moved with speed')
            self.stage_tom.Speed = sp0
        
            
    def getGunShift(self):
        x = self.gun.Shift.X
        y = self.gun.Shift.Y
        return x, y 
    
    def setGunShift(self, x, y):
        """x y can only be float numbers between -1 and 1"""
        gs = self.gun.Shift
        if abs(x) > 1 or abs(y) > 1:
            raise ValueError("GunShift x/y must be a floating number between -1 an 1. Input: x={}, y={}".format(x, y))
        
        if x is not None:
            gs.X = x
        if y is not None:
            gs.Y = y
            
        self.gun.Shift = gs
    
    def getGunTilt(self):
        x = self.gun.Tilt.X
        y = self.gun.Tilt.Y
        return x, y
    
    def setGunTilt(self, x, y):
        gt = self.tecnai.Gun.Tilt
        if abs(x) > 1 or abs(y) > 1:
            raise ValueError("GunTilt x/y must be a floating number between -1 an 1. Input: x={}, y={}".format(x, y))
        
        if x is not None:
            gt.X = x
        if y is not None:
            gt.Y = y
            
        self.tecnai.Gun.Tilt = gt
        
    def getBeamAlignShift(self):
        """Align Shift"""
        x = self.illu_tom.BeamAlignShift.X
        y = self.illu_tom.BeamAlignShift.Y
        return x, y
    
    def setBeamAlignShift(self, x, y):
        """Align Shift"""
        bs = self.illu_tom.BeamAlignShift
        if abs(x) > 1 or abs(y) > 1:
            raise ValueError("BeamAlignShift x/y must be a floating number between -1 an 1. Input: x={}, y={}".format(x, y))
            
        if x is not None:
            bs.X = x
        if y is not None:
            bs.Y = y
        self.illu_tom.BeamAlignShift = bs
        
    def getBeamTilt(self):
        """rotation center in FEI"""
        x = self.illu_tom.BeamAlignmentTilt.X
        y = self.illu_tom.BeamAlignmentTilt.Y
        return x, y
    
    def setBeamTilt(self, x, y):
        """rotation center in FEI"""
        bt = self.illu_tom.BeamAlignmentTilt
        
        if x is not None:
            if abs(x) > 1:
                raise ValueError("BeamTilt x must be a floating number between -1 an 1. Input: x={x}".format(x))
            bt.X = x
        if y is not None:
            if abs(y) > 1:
               raise ValueError("BeamTilt y must be a floating number between -1 an 1. Input: y={y}".format(y))
            bt.Y = y
        self.illu_tom.BeamAlignmentTilt = bt

    def getImageShift1(self):
        """User image shift"""
        return self.proj_tom.ImageShift.X, self.proj_tom.ImageShift.Y

    def setImageShift1(self, x, y):
        is1 = self.proj_tom.ImageShift
        if abs(x) > 1 or abs(y) > 1:
            raise ValueError("ImageShift1 x/y must be a floating number between -1 an 1. Input: x={}, y={}".format(x, y))
            
        if x is not None:
            is1.X = x
        if y is not None:
            is1.Y = y
        
        self.proj_tom.ImageShift = is1

    def getImageShift2(self):
        return self.proj_tom.ImageBeamShift.X, self.proj_tom.ImageBeamShift.Y

    def setImageShift2(self, x, y):
        is2 = self.proj_tom.ImageBeamShift
        if abs(x) > 1 or abs(y) > 1:
            raise ValueError("ImageShift2 x/y must be a floating number between -1 an 1. Input: x={}, y={}".format(x, y))
            
        if x is not None:
            is2.X = x
        if y is not None:
            is2.Y = y
        
        self.proj_tom.ImageBeamShift = is2

    def isStageMoving(self):
        if self.stage.Status == 0:
            return False
        else:
            return True
    
    def stopStage(self):
        #self.stage.Status = self.goniostopped
        raise NotImplementedError

    def getProbeMode(self):
        # would be alpha mode on a JEOL
        return PROBE_MODES[self.illu_tom.ProbeMode]
        
    def setProbeMode(self, value):
        if instance(value, int):
            self.illu_tom.ProbeMode = value
        elif isinstance(value, str):
            self.illu_tom.ProbeMode = PROBE_MODES.index(value)
            
    def getFunctionMode(self):
        """mag, D, or LAD"""
        mode = self.proj_tom.Submode
        return FUNCTION_MODES[mode]

    def setFunctionMode(self, value):
        """mag, D, or LAD"""
        if isinstance(value, str):
            try:
                value = FUNCTION_MODES.index(value)
            except ValueError:
                raise ValueError("Unrecognized function mode: {}".format(value))
        self.FunctionMode_value = value
    
    def getModeString(self):
        return self.tem.Projection.SubModeString
    
    def getHolderType(self):
        return self.stage.Holder
        
    """What is the relationship between defocus and focus?? Both are changing the defoc value"""
    def getDiffFocus(self):
        return self.proj_tom.Defocus

    def setDiffFocus(self, value):
        """defocus value in unit m"""
        self.proj_tom.Defocus = value
        
    def getFocus(self):
        return self.proj_tom.Focus
    
    def setFocus(self, value):
        self.proj_tom.Focus = value
        
    def getApertureSize(self, aperture):
         if aperture == 'C1':
             return self.illu_tom.C1ApertureSize * 1e3
         elif aperture == 'C2':
             return self.illu_tom.C2ApertureSize * 1e3
         else:
             raise ValueError("aperture must be specified as 'C1' or 'C2'.")
         
    def getBeamShift(self):
        return self.illu_tom.BeamShift.X, self.illu_tom.BeamShift.Y
    
    def setBeamShift(self, x, y):
        us = self.illu_tom.BeamShift
        if x > 0 or y > 0 or x < -1 or y < -1:
            raise ValueError("BeamShift x/y must be a floating number between -1 an 0. Input: x={}, y={}".format(x, y))
            return
        
        if x is not None:
            us.X = x
            
        if y is not None:
            us.Y = y
        
        self.illu_tom.BeamShift = us
        
    def getDarkFieldTilt(self):
        return self.illu_tom.DarkfieldTilt.X, self.illu_tom.DarkfieldTilt.Y
    
    def setDarkFieldTilt(self, x, y):
        """does not set"""
        return 0
    
    def getScreenCurrent(self):
        """return screen current in nA"""
        return self.tom.Screen.Current * 1e9
    
    def isfocusscreenin(self):
        return self.tom.Screen.IsFocusScreenIn
    
    def getScreenPosition(self):
        pos = self.tom.Screen.Position
        if pos == 0:
            return 'down'
        elif pos == 1:
            return 'up'
        elif pos == 2:
            return 'moving_down'
        elif pos == 3:
            return 'moving_up'
        
    def setScreenPosition(self, pos):
        if isinstance(pos, int):
            self.tom.Screen.SetScreenPosition(pos)
        elif isinstance(pos, str):
            if pos == 'down':
                self.tom.Screen.SetScreenPosition(0)
            elif pos == 'up':
                self.tom.Screen.SetScreenPosition(1)
    
    def getDiffShift(self):
        """To be tested"""
        if self.proj.Mode != 1:
            return (0, 0)
        
        return self.proj.DiffractionShift.X,self.proj.DiffractionShift.Y
        
    def setDiffShift(self, x, y):
        """To be tested"""
        ds = self.proj.DiffractionShift
        if x > 1 or y > 1 or x < -1 or y < -1:
            print("Invalid PLA setting: can only be float numbers between -1 and 0.")
            return
        
        if x is not None:
            ds.X = x
            
        if y is not None:
            ds.Y = y
        
        self.proj.DiffractionShift = ds

    def releaseConnection(self):
        comtypes.CoUninitialize()
        logger.info("Connection to microscope released")
        print("Connection to microscope released")

    def isBeamBlanked(self):
        """to be tested"""
        return self.illu.BeamBlanked

    def setBeamBlank(self, value):
        self.illu.BeamBlanked = value
    
    def setBeamUnblank(self):
        self.illu.BeamBlanked = 0

    def getCondensorLensStigmator(self):
        return self.illu_tom.CondenserStigmator.X, self.illu_tom.CondenserStigmator.Y

    def setCondensorLensStigmator(self, x, y):
        self.illu_tom.CondenserStigmator.X = x
        self.illu_tom.CondenserStigmator.Y = y
        
    def getIntermediateLensStigmator(self):
        """diffraction stigmator"""
        return self.illu_tom.DiffractionStigmator.X, self.illu_tom.DiffractionStigmator.Y

    def setIntermediateLensStigmator(self, x, y):
        self.illu_tom.DiffractionStigmator.X = x
        self.illu_tom.DiffractionStigmator.Y = y

    def getObjectiveLensStigmator(self):
        return self.illu_tom.ObjectiveStigmator.X, self.illu_tom.ObjectiveStigmator.Y

    def setObjectiveLensStigmator(self, x, y):
        self.illu_tom.ObjectiveStigmator.X = x
        self.illu_tom.ObjectiveStigmator.Y = y

    def getSpotSize(self):
        """0-based indexing for GetSpotSize, add 1 for consistency"""
        return self.illu_tom.SpotsizeIndex
    
    def setSpotSize(self, value):
        self.illu_tom.SpotsizeIndex = value
    
    def getMagnificationIndex(self):
        if self.proj_tom.Mode != 1:
            ind = self.proj.MagnificationIndex
            return ind
        else:
            ind = self.proj.CameraLengthIndex
            return ind

    def setMagnificationIndex(self, index):
        if self.proj_tom.Mode != 1:
            self.proj.MagnificationIndex = index
        else:
            self.proj.CameraLengthIndex = index
    
    def getBrightness(self):
        """return diameter in microns"""
        return self.illu_tom.Intensity

    def setBrightness(self, value):
        self.illu_tom.Intensity = value

    def normalizeLenses(self, what='all', lens_ID=None):
        """
        Normalize microscope lenses. Good idea to do after automated mode switches, if reproducibility is desired.
        :param what: 'all', 'illumination', 'projection', or enum ID 1-6, 11, 12, 13
        :return:
        """
        if what.lower() == 'all':
            self.illu.Normalize(6)
            self.proj.Normalize(12)
        elif what.lower() == 'illumination':
            if lens_ID is None:
                lens_ID = 6
            self.illu.Normalize(lens_ID)
        elif what.lower() == 'projection':
            if lens_ID is None:
                lens_ID = 12
            self.proj.Normalize(lens_ID)
        else:
            raise ValueError('what parameter must be illumination, projection, all')
            
    def getSTEMMagnification(self):
        return self.stem_tom.Magnification
        
    def setSTEMMagnification(self, value):
        self.stem_tom.Magnification = value
        
    def setSTEMMode(self, value):
        if value == 'nano':
            self.stem_tom.Mode = 1
            self.setProbeMode(1)
        elif value == 'micro':
            self.stem_tom.Mode = 1
            self.setProbeMode(0)
        elif value == 'LM':
            self.stem_tom.Mode = 0
        else:
            raise ValueError('STEM mode must be nano, micro, or LM')
            
    def getSTEMRotation(self):
        return self.stem_tom.Rotation / pi * 180
        
    def setSTEMRotiation(self, value):
        self.stem_tom.Rotation = value / 180 * pi
        
    def getSTEMFocusStrategy(self):
        return STEM_FOCUS_STRATEGY[self.stem_tom.FocusStrategy]
        
    def setSTEMFocusStrategy(self, value):
        if isinstance(value, int):
            self.stem_tom.FocusStrategy = value
        elif isinstance(value, str):
            self.stem_tom.FocusStrategy = STEM_FOCUS_STRATEGY.index(value)
            
    def insertHAADF(self):
        # I have found no better way of doing this... make sure that "auto 
        # insert/retract" is checked
        self.setScreenPosition('down')
        
    def retractHAADF(self):
        self.setScreenPosition('up')

    def setTIADwellTime(self, dwell_time=2e-6):
        """
        This function can be used to set the detector filter to reasonable values by taking a short dummy exposure.
        Better do NOT use this for the _scan_ filter
        :param dwell_time: Dwell time per pixel in seconds
        :return: nothing
        """
        self.acq.AddAcqDevice(self.acq.Detectors[0])
        self.acq.Detectors.AcqParams.DwellTime = dwell_time
        self.acq.Detectors.AcqParams.ImageSize = 2
        self.acq.Detectors.AcqParams.Binning = 8
        ff = self.acq.AcquireImages()

    def getStatusDict(self):
        """
        Returns a selection of TEM Metadata useful to store with results
        """
        
        ill = {'beam_blanked': self.isBeamBlanked(),
               'spot_size': self.getSpotSize(),
                'brightness': self.getBrightness(),
                'probe_mode': self.getProbeMode(),
                'beam_shift': self.getBeamShift(),
                'beam_tilt': self.getBeamTilt()}    
                
        pro = {'diffraction': self.proj.Mode == 1,
        'projection_mode': self.getFunctionMode(),
               'projection_sub_mode': self.getModeString(),
                'focus': self.getFocus(),
                'diffraction_focus': self.getDiffFocus(),
                'image_shift': self.getImageShift1(),
                'image_beam_shift': self.getImageShift2(),
                'diffraction_shift': self.getDiffShift(),
                'nominal_magnification': -1,
                'magnification_index': -1,
                'nominal_camera_length': -1,
                'nominal_camera_length_index': -1}
                
        if self.proj.mode == 0:
            pro.update({'magnification_index': self.getMagnificationIndex(), 
            'nominal_magnification': self.getMagnification()})
        elif self.proj.mode == 1:
            pro.update({'camera_length_index': self.getMagnificationIndex(), 
            'nominal_camera_length': self.getMagnification()})
        
        stg = {ax: val for ax, val in zip(['x', 'y', 'z', 'a', 'b'],
                                          self.getStagePosition())} 
                                          
        gun = {'voltage': self.getHTValue()}
        
        return {'illumination': ill,
                'projection': pro,
                'gun': gun,
                'stage': stg,
                'stem': {}}




