import os
import sys
import comtypes.client

# set the following flag to True to attach to an existing instance of the program
# otherwise a new instance of the program will be started
AttachToInstance = False

# set the following flag to True to manually specify the path to ETABS.exe
# this allows for a connection to a version of ETABS other than the latest installation
# otherwise the latest installed version of ETABS will be launched
SpecifyPath = False

# if the above flag is set to True, specify the path to ETABS below
ProgramPath = "C:\Program Files (x86)\Computers and Structures\ETABS 17\ETABS.exe"

# full path to the model
# set it to the desired path of your model
APIPath = 'C:\CSi_ETABS_API_Example'
if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass
ModelPath = APIPath + os.sep + 'API_1-001.edb'

if AttachToInstance:
    # attach to a running instance of ETABS
    try:
        # get the active ETABS object
        myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)

else:
    # create API helper object
    helper = comtypes.client.CreateObject('ETABSv17.Helper')
    helper = helper.QueryInterface(comtypes.gen.ETABSv17.cHelper)
    if SpecifyPath:
        try:
            # 'create an instance of the ETABS object from the specified path
            myETABSObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program from " + ProgramPath)
            sys.exit(-1)
    else:

        try:
            # create an instance of the ETABS object from the latest installed ETABS
            myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program.")
            sys.exit(-1)

    # start ETABS application
    myETABSObject.ApplicationStart()

# Assign model to SapModel variable
SapModel = myETABSObject.SapModel

# initialize model
SapModel.InitializeNewModel()

# create new blank model
ret = SapModel.File.NewBlank()

# Unlock model so we can edit
SapModel.SetModelIsLocked(False)

# Define units
kN_m_C = 6
ret = SapModel.SetPresentUnits(kN_m_C)
print(ret)

# Define steel material
SapModel.PropMaterial.SetMaterial("STEEL355", 1)

# Define detailed material properties
SapModel.PropMaterial.SetOSteel(
    "STEEL355",  # name of the material
    355000,  # fy, kPa
    510000,  # fu, kPa
    390500,  # Fye, kPa, effective yield stress
    561000,  # Fue, kPa, effective tensile Strength
    1,  # stress-strain curve type, 0 - user defined, 1 - parametric simple
    7,  # stress-strain curve type; 7 - isotropic
    0.015,  # StrainAtHardening
    0.11,  # StrainAtMaxStress
    0.17,  # StrainAtRupture
)

# Import and define section from etabs library
SapModel.PropFrame.ImportProp(
    "SHHF100X100X5",  # section name (any name is ok)
    "STEEL355",  # material
    "BSShapes2006.xml",  # library file (check etabs installation folder)
    "SHHF100X100X5",  # name of the section inside the library
)

spans = [2.5, 4, 4, 4, 2.5]

x_c = [0]  # x coordinate of left support is zero

for i in range(0, len(spans)):
    x_c.append(x_c[i] + spans[i])

joints = []  # initialize empty list to store name of joints

# loop to create upper and lower chord joints
for i in range(len(x_c)):
    x = x_c[i]  # x coordinate
    y = 0  # y coordinate
    z = 4  # z coordinate
    joints.append('joint_' + str(i))  # populate list with joints names
    ret = SapModel.PointObj.AddCartesian(x, y, z, '', joints[i])  # add joint and store joint name

frames = []  # initialize empty list to store name of frames

# loop to create frames
for i in range(0, len(joints) - 1):
    frames.append('frame_' + str(i))  # populate list with frame names
    ret = SapModel.FrameObj.AddByPoint(joints[i], joints[i + 1], '', 'SHHF100X100X5', frames[i])

# Define load patterns
SapModel.LoadPatterns.Add("SDL",  # name of the load pattern
                          2,  # type of load pattern (2 = SuperDead)
                          0,  # self-weight multiplier
                          True)  # static linear load case if True
SapModel.LoadPatterns.Add("LL",  # live load case
                          3,  # type of load pattern (3 = Live)
                          0,
                          True)

# Create combinations
SapModel.RespCombo.Add("COMB1 - ULS",  # load combination name
                       0)  # combo type (0=linear additive; 1=Envelope; ... )
SapModel.RespCombo.Add("COMB2 - SLS", 0)  # same as before, but for comb2

# Add cases to the combinations with partial factors
# comb1
SapModel.RespCombo.SetCaseList("COMB1 - ULS",  # name of the combo
                               0,  # case or combo (0=case; 1=combo)
                               "Dead",  # name of the load case
                               1.35)  # partial factor
SapModel.RespCombo.SetCaseList("COMB1 - ULS",  # name of the combo
                               0,  # case or combo (0=case; 1=combo)
                               "SDL",  # name of the load case
                               1.35)  # partial factor
SapModel.RespCombo.SetCaseList("COMB1 - ULS", 0, "LL", 1.5)

# comb2
SapModel.RespCombo.SetCaseList("COMB2 - SLS",  # name of the combo
                               0,  # case or combo (0=case; 1=combo)
                               "Dead",  # name of the load case
                               1.0)  # partial factor
SapModel.RespCombo.SetCaseList("COMB2 - SLS",  # name of the combo
                               0,  # add combo type
                               "SDL",  # name of the load case
                               1.0)  # partial factor
SapModel.RespCombo.SetCaseList("COMB2 - SLS", 0, "LL", 1.0)

# Loads
# define load values
sdl_load = -10  # kN/m distributed load
ll_load = -5  # kN/m distributed load

# assign load values
for frame in frames:
    ret = SapModel.FrameObj.SetLoadDistributed(frame,
                                               "SDL",
                                               1,  # 1 = force per unit length; 2 = moment/length
                                               6,  # 6 = Z direction
                                               0,  # relative distance from i-end
                                               1,  # relative distance from j-end
                                               sdl_load,  # load value at the start
                                               sdl_load,  # load value at the end
                                               "Global",  # coordinate system
                                               True,  # True=relative distance, False=absolute
                                               True,  # Replace loads = True
                                               )

# assign load values
for frame in frames:
    ret = SapModel.FrameObj.SetLoadDistributed(frame,
                                               "LL",
                                               1,  # 1 = force per unit length; 2 = moment/length
                                               6,  # 6 = Z direction
                                               0,  # relative distance from i-end
                                               1,  # relative distance from j-end
                                               ll_load,  # load value at the start
                                               ll_load,  # load value at the end
                                               "Global",  # coordinate system
                                               True,  # True=relative distance, False=absolute
                                               True,  # Replace loads = True
                                               )

# Supports
# set restraint values
fixed_support = [True, True, True, True, False, True]  # u1,u2,u3,r1,r2,r3 - fixed support
roller_support = [False, True, True, True, False, True]  # u1,u2,u3,r1,r2,r3 - rolling support

# assign restraints to left support
SapModel.PointObj.SetRestraint(joints[0], fixed_support)

# assign roller support to remaining joints
for i in range(1, len(joints)):
    SapModel.PointObj.SetRestraint(joints[i], roller_support)

# define 2D degrees of freedom
dof = [True, False, True, False, True, False]  # ux, uy, uz, rx, ry, rz
# set degrees of freedom
SapModel.Analyze.SetActiveDOF(dof)

# Save the model
ModelPath = 'C:\\Users\\Miguel.Tavares\\Desktop\\ETABS MODELS\\Python API\\ContinuousBeam.edb'
SapModel.File.Save(ModelPath)

# Run Analysis
SapModel.Analyze.RunAnalysis()

# Deselect All Load Combinations
SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput

# Select load combo
SapModel.Results.Setup.SetComboSelectedForOutput("COMB1 - ULS")

# Get results for a frame object
# Initialize dummy variables
ObjectElm = 0  # equal to zero to retrieve by name of the frame
NumberResults = 0  # The total number of results returned by the program
Obj = []  # line object name associated with each result, if any
ObjSta = []  # distance measured from the I-end of the line object to the result location
Elm = []  # line element name associated with each result
ElmSta = []  # distance measured from the I-end of the line element to the result location
LoadCase = []  # name of the analysis case or load combination
StepType = []  # step type, if any, for each result
StepNum = []  # step number, if any, for each result
P = []  # axial force for each result
V2 = []
V3 = []
T = []
M2 = []
M3 = []

# Get results for frame 2
frame2results = SapModel.Results.FrameForce('frame_2',
                                            ObjectElm,
                                            NumberResults,  # The total number of results returned by the program
                                            Obj,  # line object name associated with each result, if any
                                            ObjSta,
                                            # distance measured from the I-end of the line object to the result location
                                            Elm,  # line element name associated with each result
                                            ElmSta,
                                            # distance measured from the I-end of the line element to the result location
                                            LoadCase,  # name of the analysis case or load combination
                                            StepType,  # step type, if any, for each result
                                            StepNum,  # step number, if any, for each result
                                            P,  # axial force for each result
                                            V2,
                                            V3,
                                            T,
                                            M2,
                                            M3,
                                            )
