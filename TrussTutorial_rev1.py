import sys
import comtypes.client

# This seems to only work when pycharm is opened AFTER the etabs object
#is opened

#set the following flag to True to attach to an existing instance of the program
#otherwise a new instance of the program will be started
AttachToInstance = True

if AttachToInstance:
    #attach to a running instance of ETABS
    try:
        #get the active ETABS object
        myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)

# Get current etabs instance and assign it to ETABSObject variable
myETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")

# Assign model to SapModel variable
SapModel = myETABSObject.SapModel

# Unlock model so we can edit
SapModel.SetModelIsLocked(False)

# Define units
kN_m_C = 6
SapModel.SetPresentUnits(kN_m_C)

# Define steel material
SapModel.PropMaterial.SetMaterial("STEEL355", 1)

# Define detailed material properties
SapModel.PropMaterial.SetOSteel(
    "STEEL355", # name of the material
    355000,     # fy, kPa
    510000,     # fu, kPa
    390500,     # Fye, kPa, effective yield stress
    561000,     # Fue, kPa, effective tensile Strength
    1,          # stress-strain curve type, 0 - user defined, 1 - parametric simple
    7,          # stress-strain curve type; 7 - isotropic
    0.015,      # StrainAtHardening
    0.11,       # StrainAtMaxStress
    0.17,       # StrainAtRupture
)

# Import and define section from etabs library
SapModel.PropFrame.ImportProp(
    "SHHF100X100X5",    # section name (any name is ok)
    "STEEL355",         # material
    "BSShapes2006.xml", # library file (check etabs installation folder)
    "SHHF100X100X5",    # name of the section inside the library
)

# Define geometry variables of the truss
L = 30  # total length of the truss
h = 2   # truss depth
n = 10  # number of bays

upper_chord = []  # initialize empty list to store name of joints - upper chord
lower_chord = []  # initialize empty list to store name of joints - lower chord


# loop to create upper and lower chord joints
for i in range(n+1):
    x = i * L/n  # x coordinate
    y = 0  # y coordinate
    z = 12  # z coordinate
    upper_chord.append('upper_chord_' + str(i))  # populate list with joints names
    lower_chord.append('lower_chord_' + str(i))  # populate list with joints names
    ret = SapModel.PointObj.AddCartesian(x, y, z, '', upper_chord[i]) # add joint and store joint name
    ret = SapModel.PointObj.AddCartesian(x, y, z-h, '', lower_chord[i]) # add joint and store joint name


upper_chord_f = []  # initialize empty list to store name of frames - upper chord
lower_chord_f = []  # initialize empty list to store name of frames - lower chord
diagonal_f = [] # initialize empty list to store name of frames - diagonals

# loop to create upper chord, lower chord and diagonal frames
for i in range(n):
    upper_chord_f.append('upper_chord_f_' + str(i))  # populate list with frame names
    lower_chord_f.append('lower_chord_f_' + str(i))  # populate list with frame names
    diagonal_f.append('diagonals_f_' + str(i))  # populate list with frame names
    ret = SapModel.FrameObj.AddByPoint(upper_chord[i], upper_chord[i+1], '', 'SHHF100X100X5', upper_chord_f[i])
    ret = SapModel.FrameObj.AddByPoint(lower_chord[i], lower_chord[i+1], '', 'SHHF100X100X5', lower_chord_f[i])
    ret = SapModel.FrameObj.AddByPoint(upper_chord[i], lower_chord[i+1], '', 'SHHF100X100X5', diagonal_f[i])


post_f = []  # initialize empty list to store name of frames - posts

# loop to create post frames
for i in range(n+1):
    post_f.append('post_f_' + str(i))  # populate list with frame names
    ret = SapModel.FrameObj.AddByPoint(upper_chord[i], lower_chord[i], '', 'SHHF100X100X5', post_f[i])


# Releases
# define pinned condition with boolean arrays for both ends
release_ii = [False, False, False, True, True, True]  # set release condition for truss frames (i-end; rotations released)
release_jj = [False, False, False, False, True, True]  # set release condition for truss frames (j-end; rotations released)

# define array with spring values of stiffness for releases (pinned = 0)
start_value = [0, 0, 0, 0, 0, 0]
end_value = [0, 0, 0, 0, 0, 0]

# Loops to assign releases
for i in range(n):
    ret = SapModel.FrameObj.SetReleases(upper_chord_f[i], release_ii, release_jj, start_value, end_value)
    ret = SapModel.FrameObj.SetReleases(lower_chord_f[i], release_ii, release_jj, start_value, end_value)
    ret = SapModel.FrameObj.SetReleases(diagonal_f[i], release_ii, release_jj, start_value, end_value)

for post in post_f:
    ret = SapModel.FrameObj.SetReleases(post, release_ii, release_jj, start_value, end_value)


# Define load patterns
SapModel.LoadPatterns.Add("SDL",  # name of the load pattern
                          2,  # type of load pattern (2 = SuperDead)
                          0,  # self-weight multiplier
                          True)  # static linear load case if True
SapModel.LoadPatterns.Add("LL",  # live load case
                          3,  # type of load pattern (3 = Live)
                          0,
                          True)

# Assign static linear load cases to load patterns created (uncomment 2 lines below if this is not done above)
# SapModel.LoadCases.StaticLinear.SetCase("SDL")
# SapModel.LoadCases.StaticLinear.SetCase("LL")

# Create combinations
SapModel.RespCombo.Add("COMB1 - ULS",  # load combination name
                       0)  # combo type (0=linear additive; 1=Envelope; ... )
SapModel.RespCombo.Add("COMB2 - SLS", 0)  # same as before, but for comb2

# Add cases to the combinations with partial factors
# comb1
SapModel.RespCombo.SetCaseList("COMB1 - ULS",  # name of the combo
                               0,  # case or combo (0=case; 1=combo)
                               "SDL",  # name of the load case
                               1.35)  # partial factor
SapModel.RespCombo.SetCaseList("COMB1 - ULS", 0, "LL", 1.5)

# comb2
SapModel.RespCombo.SetCaseList("COMB2 - SLS",  # name of the combo
                               0,  # add combo type
                               "SDL",  # name of the load case
                               1.0)  # partial factor
SapModel.RespCombo.SetCaseList("COMB2 - SLS", 0, "LL", 1.0)

# Loads
# define load values
sdl_load = (0, 0, -100, 0, 0, 0)  # F1, F2, F3, M1, M2, M3
ll_load = (0, 0, -50, 0, 0, 0)  # F1, F2, F3, M1, M2, M3

# assign load values
for joint in upper_chord:
    ret = SapModel.PointObj.SetLoadForce(joint, "SDL", sdl_load, True)  # sdl pattern; True = replace loads

for joint in upper_chord:
    ret = SapModel.PointObj.SetLoadForce(joint, "LL", ll_load, True)  # ll pattern

# Supports
# left support restraint values
left_support_restraint = [True, True, True, False, False, False]  # u1,u2,u3,r1,r2,r3 - fixed support
right_support_restraint = [False, True, True, False, False, False]  # u1,u2,u3,r1,r2,r3 - rolling support

# assign restraints to left and right lower chord joints
SapModel.PointObj.SetRestraint(lower_chord[0], left_support_restraint)
SapModel.PointObj.SetRestraint(lower_chord[-1], right_support_restraint)

# Model degrees of freedom

# define 2D degrees of freedom
dof = [True, False, True, False, True, False]  # ux, uy, uz, rx, ry, rz
# set degrees of freedom
SapModel.Analyze.SetActiveDOF(dof)

# Save the model
ModelPath = 'C:\\Users\\Miguel.Tavares\\Desktop\\ETABS MODELS\\Python API\\TrussTutorial.edb'
SapModel.File.Save(ModelPath)

# Run Analysis
SapModel.Analyze.RunAnalysis()






