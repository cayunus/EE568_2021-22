' ----------------------------------------------------------------------
' Script Created by RMxprt to generate Maxwell 3D Version 2016.0 design 
' Can specify one arg to setup external circuit                         
' ----------------------------------------------------------------------

On Error Resume Next
Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
On Error Goto 0
Set oDesktop = oAnsoftApp.GetAppDesktop()
Set oArgs = AnsoftScript.arguments
oDesktop.RestoreWindow
Set oProject = oDesktop.GetActiveProject()
Set oDesign = oProject.GetActiveDesign()
designName = oDesign.GetName
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", "mm", _
  "Rescale:=", false)
oDesign.SetSolutionType "Transient"
Set oModule = oDesign.GetModule("BoundarySetup")
if (oArgs.Count = 1) then 
oModule.EditExternalCircuit oArgs(0), Array(), Array(), Array(), Array()
end if
oEditor.SetModelValidationSettings Array("NAME:Validation Options", _
  "EntityCheckLevel:=", "Strict", "IgnoreUnclassifiedObjects:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "InsulatorThreshold:=", 2.5e+06)
On Error Resume Next
Set oModule = oDesign.GetModule("MeshSetup")
oModule.InitialMeshSettings Array("NAME:MeshSettings", "MeshMethod:=", _
  "AnsoftTAU")
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:fractions", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "4"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:halfAxial", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:endRegion", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "678.85297560927211mm"))))
On Error Goto 0
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.ModifyLibraries designName, Array("NAME:PersonalLib"), _
  Array("NAME:UserLib"), Array("NAME:SystemLib", "Materials:=", Array(_
  "Materials", "RMxprt"))
if (oDefinitionManager.DoesMaterialExist("M22_24G_3DSF0.900")) then
else
oDefinitionManager.AddMaterial Array("NAME:M22_24G_3DSF0.900", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 160, "Y:=", 0.8), Array("NAME:Coordinate", "X:=", _
  240, "Y:=", 1), Array("NAME:Coordinate", "X:=", 320, "Y:=", 1.1), Array(_
  "NAME:Coordinate", "X:=", 420, "Y:=", 1.2), Array("NAME:Coordinate", "X:=", _
  680, "Y:=", 1.3), Array("NAME:Coordinate", "X:=", 810, "Y:=", 1.325), _
  Array("NAME:Coordinate", "X:=", 1100, "Y:=", 1.375), Array(_
  "NAME:Coordinate", "X:=", 1400, "Y:=", 1.4), Array("NAME:Coordinate", "X:=", _
  2600, "Y:=", 1.475), Array("NAME:Coordinate", "X:=", 4200, "Y:=", 1.54), _
  Array("NAME:Coordinate", "X:=", 5087.91, "Y:=", 1.575), Array(_
  "NAME:Coordinate", "X:=", 6579.41, "Y:=", 1.625), Array("NAME:Coordinate", _
  "X:=", 8231.91, "Y:=", 1.675), Array("NAME:Coordinate", "X:=", 10258.9, _
  "Y:=", 1.725), Array("NAME:Coordinate", "X:=", 13348.9, "Y:=", 1.775), _
  Array("NAME:Coordinate", "X:=", 18788.9, "Y:=", 1.825), Array(_
  "NAME:Coordinate", "X:=", 28908.9, "Y:=", 1.875), Array("NAME:Coordinate", _
  "X:=", 46508.9, "Y:=", 1.925), Array("NAME:Coordinate", "X:=", 74308.9, _
  "Y:=", 1.975), Array("NAME:Coordinate", "X:=", 129350, "Y:=", 2.05), Array(_
  "NAME:Coordinate", "X:=", 248716, "Y:=", 2.2), Array("NAME:Coordinate", _
  "X:=", 447658, "Y:=", 2.45), Array("NAME:Coordinate", "X:=", 726178, "Y:=", _
  2.8), Array("NAME:Coordinate", "X:=", 3.51138e+06, "Y:=", 6.3))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 177.7, "core_loss_kc:=", 1.33, _
  "core_loss_ke:=", 1.22, "conductivity:=", 2e+06, "mass_density:=", 7650, _
  Array("NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.9", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
end if
if (oDefinitionManager.DoesMaterialExist("M22_24G_3DSF0.960")) then
else
oDefinitionManager.AddMaterial Array("NAME:M22_24G_3DSF0.960", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 160, "Y:=", 0.8), Array("NAME:Coordinate", "X:=", _
  240, "Y:=", 1), Array("NAME:Coordinate", "X:=", 320, "Y:=", 1.1), Array(_
  "NAME:Coordinate", "X:=", 420, "Y:=", 1.2), Array("NAME:Coordinate", "X:=", _
  680, "Y:=", 1.3), Array("NAME:Coordinate", "X:=", 810, "Y:=", 1.325), _
  Array("NAME:Coordinate", "X:=", 1100, "Y:=", 1.375), Array(_
  "NAME:Coordinate", "X:=", 1400, "Y:=", 1.4), Array("NAME:Coordinate", "X:=", _
  2600, "Y:=", 1.475), Array("NAME:Coordinate", "X:=", 4200, "Y:=", 1.54), _
  Array("NAME:Coordinate", "X:=", 5087.91, "Y:=", 1.575), Array(_
  "NAME:Coordinate", "X:=", 6579.41, "Y:=", 1.625), Array("NAME:Coordinate", _
  "X:=", 8231.91, "Y:=", 1.675), Array("NAME:Coordinate", "X:=", 10258.9, _
  "Y:=", 1.725), Array("NAME:Coordinate", "X:=", 13348.9, "Y:=", 1.775), _
  Array("NAME:Coordinate", "X:=", 18788.9, "Y:=", 1.825), Array(_
  "NAME:Coordinate", "X:=", 28908.9, "Y:=", 1.875), Array("NAME:Coordinate", _
  "X:=", 46508.9, "Y:=", 1.925), Array("NAME:Coordinate", "X:=", 74308.9, _
  "Y:=", 1.975), Array("NAME:Coordinate", "X:=", 129350, "Y:=", 2.05), Array(_
  "NAME:Coordinate", "X:=", 248716, "Y:=", 2.2), Array("NAME:Coordinate", _
  "X:=", 447658, "Y:=", 2.45), Array("NAME:Coordinate", "X:=", 726178, "Y:=", _
  2.8), Array("NAME:Coordinate", "X:=", 3.51138e+06, "Y:=", 6.3))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 177.7, "core_loss_kc:=", 1.33, _
  "core_loss_ke:=", 1.22, "conductivity:=", 2e+06, "mass_density:=", 7650, _
  Array("NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.96", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
end if
if (oDefinitionManager.DoesMaterialExist("copper_100C")) then
else
oDefinitionManager.AddMaterial Array("NAME:copper_100C", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), "conductivity:=", "4.25672e+07")
end if
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5414.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Band", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Band:CreateUserDefinedPart:1", _
  "Length", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5414.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "Shaft", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "Length", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5414.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "4"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "OuterRegion", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Length", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Fractions", "fractions"
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "HalfAxial", "halfAxial"
On Error Goto 0
oEditor.Copy Array("NAME:Selections", "Selections:=", "OuterRegion")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "2"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "MasterSheet"))))
oEditor.Copy Array("NAME:Selections", "Selections:=", "MasterSheet")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "MasterSheet1:CreateUserDefinedPart:1", "InfoCore", "3"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "MasterSheet1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "SlaveSheet"))))
oEditor.Copy Array("NAME:Selections", "Selections:=", "OuterRegion")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "1"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "Tool"))))
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.SetSymmetryMultiplier "fractions*(1+halfAxial)"
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignMaster Array("NAME:Master", "Objects:=", Array("MasterSheet"), _
  Array("NAME:CoordSysVector", "Origin:=", Array("0mm", "0mm", "0mm"), _
  "UPos:=", Array("2987mm", "0mm", "0mm")), "ReverseV:=", true)
oModule.AssignSlave Array("NAME:Slave", "Objects:=", Array("SlaveSheet"), _
  Array("NAME:CoordSysVector", "Origin:=", Array("0mm", "0mm", "0mm"), _
  "UPos:=", Array("1.8290099945265721e-13mm", "2987mm", "0mm")), "ReverseV:=", _
  true, "Master:=", "Master", "RelationIsSame:=", true)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/VentSlotCore", "Version:=", "12.1", "NoOfParameters:=", _
  27, "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1066mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array(_
  "NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "4.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "HalfSlot", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "VentHoles", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HoleDiaIn", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "HoleDiaOut", "Value:=", _
  "0mm"), Array("NAME:Pair", "Name:=", "HoleLocIn", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "HoleLocOut", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "VentDucts", "Value:=", "10"), Array("NAME:Pair", "Name:=", _
  "DuctWidth", "Value:=", "9.5mm"), Array("NAME:Pair", "Name:=", "DuctPitch", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", _
  "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", _
  "0"))), Array("NAME:Attributes", "Name:=", "Stator", "Flags:=", "", _
  "Color:=", "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "M22_24G_3DSF0.900", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:delta", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "-25.908893038327076deg"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:conds", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:R1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.013834093791735462ohm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:Le1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.00046156375900369416H"))))
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1066mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array(_
  "NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "4.2mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", _
  "Name:=", "CoilPitch", "Value:=", "9"), Array("NAME:Pair", "Name:=", _
  "EndExt", "Value:=", "150mm"), Array("NAME:Pair", "Name:=", "SpanExt", _
  "Value:=", "0.1mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "10deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", _
  "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", _
  "1"))), Array("NAME:Attributes", "Name:=", "Coil", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Coil"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "1.2deg", _
  "NumClones:=", "300"), Array("NAME:Options", "DuplicateBoundaries:=", _
  false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Coil"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Coil_0"))))
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5431mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "5974mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1066mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", _
  "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "300"), Array(_
  "NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "4.2mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "120mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "22.7mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "22.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", _
  "Name:=", "CoilPitch", "Value:=", "9"), Array("NAME:Pair", "Name:=", _
  "EndExt", "Value:=", "150mm"), Array("NAME:Pair", "Name:=", "SpanExt", _
  "Value:=", "0.1mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "10deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", _
  "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", _
  "2"))), Array("NAME:Attributes", "Name:=", "CoilTerm", "Flags:=", "", _
  "Color:=", "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "CoilTerm:CreateUserDefinedPart:1", "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", _
  "CoilTerm"), Array("NAME:DuplicateAroundAxisParameters", _
  "CoordinateSystemID:=", -1, "CreateNewObjects:=", true, "WhichAxis:=", "Z", _
  "AngleStr:=", "1.2deg", "NumClones:=", "300"), Array("NAME:Options", _
  "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "CoilTerm"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "CoilTerm_0"))))
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseA", "Type:=", "External", _
  "IsSolid:=", false, "Current:=", "0A", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseB", "Type:=", "External", _
  "IsSolid:=", false, "Current:=", "0A", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseC", "Type:=", "External", _
  "IsSolid:=", false, "Current:=", "0A", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
oModule.AssignCoilTerminal Array("NAME:PhA_0", "Objects:=", Array("CoilTerm_0"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_1", "Objects:=", Array("CoilTerm_1"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_2", "Objects:=", Array("CoilTerm_2"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_3", "Objects:=", Array("CoilTerm_3"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_4", "Objects:=", Array(_
  "CoilTerm_4"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_5", "Objects:=", Array(_
  "CoilTerm_5"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_6", "Objects:=", Array(_
  "CoilTerm_6"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_7", "Objects:=", Array("CoilTerm_7"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_8", "Objects:=", Array("CoilTerm_8"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_9", "Objects:=", Array("CoilTerm_9"), _
  "Conductor number:=", "conds", "Point out of terminal:=", false, _
  "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_10", "Objects:=", Array(_
  "CoilTerm_10"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_11", "Objects:=", Array(_
  "CoilTerm_11"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_12", "Objects:=", Array(_
  "CoilTerm_12"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_13", "Objects:=", Array(_
  "CoilTerm_13"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_14", "Objects:=", Array(_
  "CoilTerm_14"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_15", "Objects:=", Array(_
  "CoilTerm_15"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_16", "Objects:=", Array(_
  "CoilTerm_16"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_17", "Objects:=", Array(_
  "CoilTerm_17"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_18", "Objects:=", Array(_
  "CoilTerm_18"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_19", "Objects:=", Array(_
  "CoilTerm_19"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_20", "Objects:=", Array(_
  "CoilTerm_20"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_21", "Objects:=", Array(_
  "CoilTerm_21"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_22", "Objects:=", Array(_
  "CoilTerm_22"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_23", "Objects:=", Array(_
  "CoilTerm_23"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_24", "Objects:=", Array(_
  "CoilTerm_24"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_25", "Objects:=", Array(_
  "CoilTerm_25"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_26", "Objects:=", Array(_
  "CoilTerm_26"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_27", "Objects:=", Array(_
  "CoilTerm_27"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_28", "Objects:=", Array(_
  "CoilTerm_28"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_29", "Objects:=", Array(_
  "CoilTerm_29"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_30", "Objects:=", Array(_
  "CoilTerm_30"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_31", "Objects:=", Array(_
  "CoilTerm_31"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_32", "Objects:=", Array(_
  "CoilTerm_32"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_33", "Objects:=", Array(_
  "CoilTerm_33"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_34", "Objects:=", Array(_
  "CoilTerm_34"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_35", "Objects:=", Array(_
  "CoilTerm_35"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_36", "Objects:=", Array(_
  "CoilTerm_36"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_37", "Objects:=", Array(_
  "CoilTerm_37"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_38", "Objects:=", Array(_
  "CoilTerm_38"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_39", "Objects:=", Array(_
  "CoilTerm_39"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_40", "Objects:=", Array(_
  "CoilTerm_40"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_41", "Objects:=", Array(_
  "CoilTerm_41"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_42", "Objects:=", Array(_
  "CoilTerm_42"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_43", "Objects:=", Array(_
  "CoilTerm_43"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_44", "Objects:=", Array(_
  "CoilTerm_44"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_45", "Objects:=", Array(_
  "CoilTerm_45"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_46", "Objects:=", Array(_
  "CoilTerm_46"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_47", "Objects:=", Array(_
  "CoilTerm_47"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_48", "Objects:=", Array(_
  "CoilTerm_48"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_49", "Objects:=", Array(_
  "CoilTerm_49"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_50", "Objects:=", Array(_
  "CoilTerm_50"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_51", "Objects:=", Array(_
  "CoilTerm_51"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_52", "Objects:=", Array(_
  "CoilTerm_52"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_53", "Objects:=", Array(_
  "CoilTerm_53"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_54", "Objects:=", Array(_
  "CoilTerm_54"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_55", "Objects:=", Array(_
  "CoilTerm_55"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_56", "Objects:=", Array(_
  "CoilTerm_56"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_57", "Objects:=", Array(_
  "CoilTerm_57"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_58", "Objects:=", Array(_
  "CoilTerm_58"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_59", "Objects:=", Array(_
  "CoilTerm_59"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_60", "Objects:=", Array(_
  "CoilTerm_60"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_61", "Objects:=", Array(_
  "CoilTerm_61"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_62", "Objects:=", Array(_
  "CoilTerm_62"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_63", "Objects:=", Array(_
  "CoilTerm_63"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_64", "Objects:=", Array(_
  "CoilTerm_64"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_65", "Objects:=", Array(_
  "CoilTerm_65"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_66", "Objects:=", Array(_
  "CoilTerm_66"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_67", "Objects:=", Array(_
  "CoilTerm_67"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_68", "Objects:=", Array(_
  "CoilTerm_68"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_69", "Objects:=", Array(_
  "CoilTerm_69"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_70", "Objects:=", Array(_
  "CoilTerm_70"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_71", "Objects:=", Array(_
  "CoilTerm_71"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_72", "Objects:=", Array(_
  "CoilTerm_72"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_73", "Objects:=", Array(_
  "CoilTerm_73"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_74", "Objects:=", Array(_
  "CoilTerm_74"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_75", "Objects:=", Array(_
  "CoilTerm_75"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_76", "Objects:=", Array(_
  "CoilTerm_76"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_77", "Objects:=", Array(_
  "CoilTerm_77"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_78", "Objects:=", Array(_
  "CoilTerm_78"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_79", "Objects:=", Array(_
  "CoilTerm_79"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_80", "Objects:=", Array(_
  "CoilTerm_80"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_81", "Objects:=", Array(_
  "CoilTerm_81"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_82", "Objects:=", Array(_
  "CoilTerm_82"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_83", "Objects:=", Array(_
  "CoilTerm_83"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_84", "Objects:=", Array(_
  "CoilTerm_84"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_85", "Objects:=", Array(_
  "CoilTerm_85"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_86", "Objects:=", Array(_
  "CoilTerm_86"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_87", "Objects:=", Array(_
  "CoilTerm_87"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_88", "Objects:=", Array(_
  "CoilTerm_88"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_89", "Objects:=", Array(_
  "CoilTerm_89"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_90", "Objects:=", Array(_
  "CoilTerm_90"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_91", "Objects:=", Array(_
  "CoilTerm_91"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_92", "Objects:=", Array(_
  "CoilTerm_92"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_93", "Objects:=", Array(_
  "CoilTerm_93"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_94", "Objects:=", Array(_
  "CoilTerm_94"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_95", "Objects:=", Array(_
  "CoilTerm_95"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_96", "Objects:=", Array(_
  "CoilTerm_96"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_97", "Objects:=", Array(_
  "CoilTerm_97"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_98", "Objects:=", Array(_
  "CoilTerm_98"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_99", "Objects:=", Array(_
  "CoilTerm_99"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_100", "Objects:=", Array(_
  "CoilTerm_100"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_101", "Objects:=", Array(_
  "CoilTerm_101"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_102", "Objects:=", Array(_
  "CoilTerm_102"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_103", "Objects:=", Array(_
  "CoilTerm_103"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_104", "Objects:=", Array(_
  "CoilTerm_104"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_105", "Objects:=", Array(_
  "CoilTerm_105"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_106", "Objects:=", Array(_
  "CoilTerm_106"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_107", "Objects:=", Array(_
  "CoilTerm_107"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_108", "Objects:=", Array(_
  "CoilTerm_108"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_109", "Objects:=", Array(_
  "CoilTerm_109"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_110", "Objects:=", Array(_
  "CoilTerm_110"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_111", "Objects:=", Array(_
  "CoilTerm_111"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_112", "Objects:=", Array(_
  "CoilTerm_112"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_113", "Objects:=", Array(_
  "CoilTerm_113"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_114", "Objects:=", Array(_
  "CoilTerm_114"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_115", "Objects:=", Array(_
  "CoilTerm_115"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_116", "Objects:=", Array(_
  "CoilTerm_116"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_117", "Objects:=", Array(_
  "CoilTerm_117"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_118", "Objects:=", Array(_
  "CoilTerm_118"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_119", "Objects:=", Array(_
  "CoilTerm_119"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_120", "Objects:=", Array(_
  "CoilTerm_120"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_121", "Objects:=", Array(_
  "CoilTerm_121"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_122", "Objects:=", Array(_
  "CoilTerm_122"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_123", "Objects:=", Array(_
  "CoilTerm_123"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_124", "Objects:=", Array(_
  "CoilTerm_124"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_125", "Objects:=", Array(_
  "CoilTerm_125"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_126", "Objects:=", Array(_
  "CoilTerm_126"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_127", "Objects:=", Array(_
  "CoilTerm_127"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_128", "Objects:=", Array(_
  "CoilTerm_128"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_129", "Objects:=", Array(_
  "CoilTerm_129"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_130", "Objects:=", Array(_
  "CoilTerm_130"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_131", "Objects:=", Array(_
  "CoilTerm_131"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_132", "Objects:=", Array(_
  "CoilTerm_132"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_133", "Objects:=", Array(_
  "CoilTerm_133"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_134", "Objects:=", Array(_
  "CoilTerm_134"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_135", "Objects:=", Array(_
  "CoilTerm_135"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_136", "Objects:=", Array(_
  "CoilTerm_136"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_137", "Objects:=", Array(_
  "CoilTerm_137"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_138", "Objects:=", Array(_
  "CoilTerm_138"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_139", "Objects:=", Array(_
  "CoilTerm_139"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_140", "Objects:=", Array(_
  "CoilTerm_140"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_141", "Objects:=", Array(_
  "CoilTerm_141"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_142", "Objects:=", Array(_
  "CoilTerm_142"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_143", "Objects:=", Array(_
  "CoilTerm_143"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_144", "Objects:=", Array(_
  "CoilTerm_144"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_145", "Objects:=", Array(_
  "CoilTerm_145"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_146", "Objects:=", Array(_
  "CoilTerm_146"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_147", "Objects:=", Array(_
  "CoilTerm_147"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_148", "Objects:=", Array(_
  "CoilTerm_148"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_149", "Objects:=", Array(_
  "CoilTerm_149"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_150", "Objects:=", Array(_
  "CoilTerm_150"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_151", "Objects:=", Array(_
  "CoilTerm_151"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_152", "Objects:=", Array(_
  "CoilTerm_152"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_153", "Objects:=", Array(_
  "CoilTerm_153"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_154", "Objects:=", Array(_
  "CoilTerm_154"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_155", "Objects:=", Array(_
  "CoilTerm_155"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_156", "Objects:=", Array(_
  "CoilTerm_156"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_157", "Objects:=", Array(_
  "CoilTerm_157"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_158", "Objects:=", Array(_
  "CoilTerm_158"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_159", "Objects:=", Array(_
  "CoilTerm_159"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_160", "Objects:=", Array(_
  "CoilTerm_160"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_161", "Objects:=", Array(_
  "CoilTerm_161"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_162", "Objects:=", Array(_
  "CoilTerm_162"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_163", "Objects:=", Array(_
  "CoilTerm_163"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_164", "Objects:=", Array(_
  "CoilTerm_164"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_165", "Objects:=", Array(_
  "CoilTerm_165"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_166", "Objects:=", Array(_
  "CoilTerm_166"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_167", "Objects:=", Array(_
  "CoilTerm_167"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_168", "Objects:=", Array(_
  "CoilTerm_168"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_169", "Objects:=", Array(_
  "CoilTerm_169"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_170", "Objects:=", Array(_
  "CoilTerm_170"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_171", "Objects:=", Array(_
  "CoilTerm_171"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_172", "Objects:=", Array(_
  "CoilTerm_172"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_173", "Objects:=", Array(_
  "CoilTerm_173"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_174", "Objects:=", Array(_
  "CoilTerm_174"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_175", "Objects:=", Array(_
  "CoilTerm_175"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_176", "Objects:=", Array(_
  "CoilTerm_176"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_177", "Objects:=", Array(_
  "CoilTerm_177"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_178", "Objects:=", Array(_
  "CoilTerm_178"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_179", "Objects:=", Array(_
  "CoilTerm_179"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_180", "Objects:=", Array(_
  "CoilTerm_180"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_181", "Objects:=", Array(_
  "CoilTerm_181"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_182", "Objects:=", Array(_
  "CoilTerm_182"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_183", "Objects:=", Array(_
  "CoilTerm_183"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_184", "Objects:=", Array(_
  "CoilTerm_184"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_185", "Objects:=", Array(_
  "CoilTerm_185"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_186", "Objects:=", Array(_
  "CoilTerm_186"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_187", "Objects:=", Array(_
  "CoilTerm_187"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_188", "Objects:=", Array(_
  "CoilTerm_188"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_189", "Objects:=", Array(_
  "CoilTerm_189"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_190", "Objects:=", Array(_
  "CoilTerm_190"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_191", "Objects:=", Array(_
  "CoilTerm_191"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_192", "Objects:=", Array(_
  "CoilTerm_192"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_193", "Objects:=", Array(_
  "CoilTerm_193"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_194", "Objects:=", Array(_
  "CoilTerm_194"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_195", "Objects:=", Array(_
  "CoilTerm_195"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_196", "Objects:=", Array(_
  "CoilTerm_196"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_197", "Objects:=", Array(_
  "CoilTerm_197"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_198", "Objects:=", Array(_
  "CoilTerm_198"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_199", "Objects:=", Array(_
  "CoilTerm_199"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_200", "Objects:=", Array(_
  "CoilTerm_200"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_201", "Objects:=", Array(_
  "CoilTerm_201"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_202", "Objects:=", Array(_
  "CoilTerm_202"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_203", "Objects:=", Array(_
  "CoilTerm_203"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_204", "Objects:=", Array(_
  "CoilTerm_204"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_205", "Objects:=", Array(_
  "CoilTerm_205"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_206", "Objects:=", Array(_
  "CoilTerm_206"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_207", "Objects:=", Array(_
  "CoilTerm_207"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_208", "Objects:=", Array(_
  "CoilTerm_208"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_209", "Objects:=", Array(_
  "CoilTerm_209"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_210", "Objects:=", Array(_
  "CoilTerm_210"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_211", "Objects:=", Array(_
  "CoilTerm_211"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_212", "Objects:=", Array(_
  "CoilTerm_212"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_213", "Objects:=", Array(_
  "CoilTerm_213"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_214", "Objects:=", Array(_
  "CoilTerm_214"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_215", "Objects:=", Array(_
  "CoilTerm_215"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_216", "Objects:=", Array(_
  "CoilTerm_216"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_217", "Objects:=", Array(_
  "CoilTerm_217"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_218", "Objects:=", Array(_
  "CoilTerm_218"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_219", "Objects:=", Array(_
  "CoilTerm_219"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_220", "Objects:=", Array(_
  "CoilTerm_220"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_221", "Objects:=", Array(_
  "CoilTerm_221"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_222", "Objects:=", Array(_
  "CoilTerm_222"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_223", "Objects:=", Array(_
  "CoilTerm_223"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_224", "Objects:=", Array(_
  "CoilTerm_224"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_225", "Objects:=", Array(_
  "CoilTerm_225"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_226", "Objects:=", Array(_
  "CoilTerm_226"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_227", "Objects:=", Array(_
  "CoilTerm_227"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_228", "Objects:=", Array(_
  "CoilTerm_228"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_229", "Objects:=", Array(_
  "CoilTerm_229"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_230", "Objects:=", Array(_
  "CoilTerm_230"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_231", "Objects:=", Array(_
  "CoilTerm_231"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_232", "Objects:=", Array(_
  "CoilTerm_232"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_233", "Objects:=", Array(_
  "CoilTerm_233"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_234", "Objects:=", Array(_
  "CoilTerm_234"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_235", "Objects:=", Array(_
  "CoilTerm_235"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_236", "Objects:=", Array(_
  "CoilTerm_236"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_237", "Objects:=", Array(_
  "CoilTerm_237"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_238", "Objects:=", Array(_
  "CoilTerm_238"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_239", "Objects:=", Array(_
  "CoilTerm_239"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_240", "Objects:=", Array(_
  "CoilTerm_240"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_241", "Objects:=", Array(_
  "CoilTerm_241"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_242", "Objects:=", Array(_
  "CoilTerm_242"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_243", "Objects:=", Array(_
  "CoilTerm_243"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_244", "Objects:=", Array(_
  "CoilTerm_244"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_245", "Objects:=", Array(_
  "CoilTerm_245"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_246", "Objects:=", Array(_
  "CoilTerm_246"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_247", "Objects:=", Array(_
  "CoilTerm_247"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_248", "Objects:=", Array(_
  "CoilTerm_248"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_249", "Objects:=", Array(_
  "CoilTerm_249"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_250", "Objects:=", Array(_
  "CoilTerm_250"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_251", "Objects:=", Array(_
  "CoilTerm_251"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_252", "Objects:=", Array(_
  "CoilTerm_252"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_253", "Objects:=", Array(_
  "CoilTerm_253"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_254", "Objects:=", Array(_
  "CoilTerm_254"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_255", "Objects:=", Array(_
  "CoilTerm_255"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_256", "Objects:=", Array(_
  "CoilTerm_256"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_257", "Objects:=", Array(_
  "CoilTerm_257"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_258", "Objects:=", Array(_
  "CoilTerm_258"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_259", "Objects:=", Array(_
  "CoilTerm_259"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_260", "Objects:=", Array(_
  "CoilTerm_260"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_261", "Objects:=", Array(_
  "CoilTerm_261"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_262", "Objects:=", Array(_
  "CoilTerm_262"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_263", "Objects:=", Array(_
  "CoilTerm_263"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_264", "Objects:=", Array(_
  "CoilTerm_264"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_265", "Objects:=", Array(_
  "CoilTerm_265"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_266", "Objects:=", Array(_
  "CoilTerm_266"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_267", "Objects:=", Array(_
  "CoilTerm_267"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_268", "Objects:=", Array(_
  "CoilTerm_268"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_269", "Objects:=", Array(_
  "CoilTerm_269"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_270", "Objects:=", Array(_
  "CoilTerm_270"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_271", "Objects:=", Array(_
  "CoilTerm_271"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_272", "Objects:=", Array(_
  "CoilTerm_272"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_273", "Objects:=", Array(_
  "CoilTerm_273"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_274", "Objects:=", Array(_
  "CoilTerm_274"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_275", "Objects:=", Array(_
  "CoilTerm_275"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_276", "Objects:=", Array(_
  "CoilTerm_276"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_277", "Objects:=", Array(_
  "CoilTerm_277"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_278", "Objects:=", Array(_
  "CoilTerm_278"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_279", "Objects:=", Array(_
  "CoilTerm_279"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_280", "Objects:=", Array(_
  "CoilTerm_280"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_281", "Objects:=", Array(_
  "CoilTerm_281"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhA_282", "Objects:=", Array(_
  "CoilTerm_282"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_283", "Objects:=", Array(_
  "CoilTerm_283"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhA_284", "Objects:=", Array(_
  "CoilTerm_284"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhCRe_285", "Objects:=", Array(_
  "CoilTerm_285"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_286", "Objects:=", Array(_
  "CoilTerm_286"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhCRe_287", "Objects:=", Array(_
  "CoilTerm_287"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhB_288", "Objects:=", Array(_
  "CoilTerm_288"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_289", "Objects:=", Array(_
  "CoilTerm_289"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhB_290", "Objects:=", Array(_
  "CoilTerm_290"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhARe_291", "Objects:=", Array(_
  "CoilTerm_291"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_292", "Objects:=", Array(_
  "CoilTerm_292"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhARe_293", "Objects:=", Array(_
  "CoilTerm_293"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseA")
oModule.AssignCoilTerminal Array("NAME:PhC_294", "Objects:=", Array(_
  "CoilTerm_294"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_295", "Objects:=", Array(_
  "CoilTerm_295"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhC_296", "Objects:=", Array(_
  "CoilTerm_296"), "Conductor number:=", "conds", "Point out of terminal:=", _
  false, "Winding:=", "PhaseC")
oModule.AssignCoilTerminal Array("NAME:PhBRe_297", "Objects:=", Array(_
  "CoilTerm_297"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_298", "Objects:=", Array(_
  "CoilTerm_298"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
oModule.AssignCoilTerminal Array("NAME:PhBRe_299", "Objects:=", Array(_
  "CoilTerm_299"), "Conductor number:=", "conds", "Point out of terminal:=", _
  true, "Winding:=", "PhaseB")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Coil", "RefineInside:=", true, _
  "Objects:=", Array("Coil_0", "Coil_1", "Coil_2", "Coil_3", "Coil_4", _
  "Coil_5", "Coil_6", "Coil_7", "Coil_8", "Coil_9", "Coil_10", "Coil_11", _
  "Coil_12", "Coil_13", "Coil_14", "Coil_15", "Coil_16", "Coil_17", "Coil_18", _
  "Coil_19", "Coil_20", "Coil_21", "Coil_22", "Coil_23", "Coil_24", "Coil_25", _
  "Coil_26", "Coil_27", "Coil_28", "Coil_29", "Coil_30", "Coil_31", "Coil_32", _
  "Coil_33", "Coil_34", "Coil_35", "Coil_36", "Coil_37", "Coil_38", "Coil_39", _
  "Coil_40", "Coil_41", "Coil_42", "Coil_43", "Coil_44", "Coil_45", "Coil_46", _
  "Coil_47", "Coil_48", "Coil_49", "Coil_50", "Coil_51", "Coil_52", "Coil_53", _
  "Coil_54", "Coil_55", "Coil_56", "Coil_57", "Coil_58", "Coil_59", "Coil_60", _
  "Coil_61", "Coil_62", "Coil_63", "Coil_64", "Coil_65", "Coil_66", "Coil_67", _
  "Coil_68", "Coil_69", "Coil_70", "Coil_71", "Coil_72", "Coil_73", "Coil_74", _
  "Coil_75", "Coil_76", "Coil_77", "Coil_78", "Coil_79", "Coil_80", "Coil_81", _
  "Coil_82", "Coil_83", "Coil_84", "Coil_85", "Coil_86", "Coil_87", "Coil_88", _
  "Coil_89", "Coil_90", "Coil_91", "Coil_92", "Coil_93", "Coil_94", "Coil_95", _
  "Coil_96", "Coil_97", "Coil_98", "Coil_99", "Coil_100", "Coil_101", _
  "Coil_102", "Coil_103", "Coil_104", "Coil_105", "Coil_106", "Coil_107", _
  "Coil_108", "Coil_109", "Coil_110", "Coil_111", "Coil_112", "Coil_113", _
  "Coil_114", "Coil_115", "Coil_116", "Coil_117", "Coil_118", "Coil_119", _
  "Coil_120", "Coil_121", "Coil_122", "Coil_123", "Coil_124", "Coil_125", _
  "Coil_126", "Coil_127", "Coil_128", "Coil_129", "Coil_130", "Coil_131", _
  "Coil_132", "Coil_133", "Coil_134", "Coil_135", "Coil_136", "Coil_137", _
  "Coil_138", "Coil_139", "Coil_140", "Coil_141", "Coil_142", "Coil_143", _
  "Coil_144", "Coil_145", "Coil_146", "Coil_147", "Coil_148", "Coil_149", _
  "Coil_150", "Coil_151", "Coil_152", "Coil_153", "Coil_154", "Coil_155", _
  "Coil_156", "Coil_157", "Coil_158", "Coil_159", "Coil_160", "Coil_161", _
  "Coil_162", "Coil_163", "Coil_164", "Coil_165", "Coil_166", "Coil_167", _
  "Coil_168", "Coil_169", "Coil_170", "Coil_171", "Coil_172", "Coil_173", _
  "Coil_174", "Coil_175", "Coil_176", "Coil_177", "Coil_178", "Coil_179", _
  "Coil_180", "Coil_181", "Coil_182", "Coil_183", "Coil_184", "Coil_185", _
  "Coil_186", "Coil_187", "Coil_188", "Coil_189", "Coil_190", "Coil_191", _
  "Coil_192", "Coil_193", "Coil_194", "Coil_195", "Coil_196", "Coil_197", _
  "Coil_198", "Coil_199", "Coil_200", "Coil_201", "Coil_202", "Coil_203", _
  "Coil_204", "Coil_205", "Coil_206", "Coil_207", "Coil_208", "Coil_209", _
  "Coil_210", "Coil_211", "Coil_212", "Coil_213", "Coil_214", "Coil_215", _
  "Coil_216", "Coil_217", "Coil_218", "Coil_219", "Coil_220", "Coil_221", _
  "Coil_222", "Coil_223", "Coil_224", "Coil_225", "Coil_226", "Coil_227", _
  "Coil_228", "Coil_229", "Coil_230", "Coil_231", "Coil_232", "Coil_233", _
  "Coil_234", "Coil_235", "Coil_236", "Coil_237", "Coil_238", "Coil_239", _
  "Coil_240", "Coil_241", "Coil_242", "Coil_243", "Coil_244", "Coil_245", _
  "Coil_246", "Coil_247", "Coil_248", "Coil_249", "Coil_250", "Coil_251", _
  "Coil_252", "Coil_253", "Coil_254", "Coil_255", "Coil_256", "Coil_257", _
  "Coil_258", "Coil_259", "Coil_260", "Coil_261", "Coil_262", "Coil_263", _
  "Coil_264", "Coil_265", "Coil_266", "Coil_267", "Coil_268", "Coil_269", _
  "Coil_270", "Coil_271", "Coil_272", "Coil_273", "Coil_274", "Coil_275", _
  "Coil_276", "Coil_277", "Coil_278", "Coil_279", "Coil_280", "Coil_281", _
  "Coil_282", "Coil_283", "Coil_284", "Coil_285", "Coil_286", "Coil_287", _
  "Coil_288", "Coil_289", "Coil_290", "Coil_291", "Coil_292", "Coil_293", _
  "Coil_294", "Coil_295", "Coil_296", "Coil_297", "Coil_298", "Coil_299"), _
  "RestrictElem:=", false, "NumMaxElem:=", "1000", "RestrictLength:=", true, _
  "MaxLength:=", "45.4mm")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignTrueSurfOp Array("NAME:SurfApprox_Coil", "Objects:=", Array(_
  "Coil_0", "Coil_1", "Coil_2", "Coil_3", "Coil_4", "Coil_5", "Coil_6", _
  "Coil_7", "Coil_8", "Coil_9", "Coil_10", "Coil_11", "Coil_12", "Coil_13", _
  "Coil_14", "Coil_15", "Coil_16", "Coil_17", "Coil_18", "Coil_19", "Coil_20", _
  "Coil_21", "Coil_22", "Coil_23", "Coil_24", "Coil_25", "Coil_26", "Coil_27", _
  "Coil_28", "Coil_29", "Coil_30", "Coil_31", "Coil_32", "Coil_33", "Coil_34", _
  "Coil_35", "Coil_36", "Coil_37", "Coil_38", "Coil_39", "Coil_40", "Coil_41", _
  "Coil_42", "Coil_43", "Coil_44", "Coil_45", "Coil_46", "Coil_47", "Coil_48", _
  "Coil_49", "Coil_50", "Coil_51", "Coil_52", "Coil_53", "Coil_54", "Coil_55", _
  "Coil_56", "Coil_57", "Coil_58", "Coil_59", "Coil_60", "Coil_61", "Coil_62", _
  "Coil_63", "Coil_64", "Coil_65", "Coil_66", "Coil_67", "Coil_68", "Coil_69", _
  "Coil_70", "Coil_71", "Coil_72", "Coil_73", "Coil_74", "Coil_75", "Coil_76", _
  "Coil_77", "Coil_78", "Coil_79", "Coil_80", "Coil_81", "Coil_82", "Coil_83", _
  "Coil_84", "Coil_85", "Coil_86", "Coil_87", "Coil_88", "Coil_89", "Coil_90", _
  "Coil_91", "Coil_92", "Coil_93", "Coil_94", "Coil_95", "Coil_96", "Coil_97", _
  "Coil_98", "Coil_99", "Coil_100", "Coil_101", "Coil_102", "Coil_103", _
  "Coil_104", "Coil_105", "Coil_106", "Coil_107", "Coil_108", "Coil_109", _
  "Coil_110", "Coil_111", "Coil_112", "Coil_113", "Coil_114", "Coil_115", _
  "Coil_116", "Coil_117", "Coil_118", "Coil_119", "Coil_120", "Coil_121", _
  "Coil_122", "Coil_123", "Coil_124", "Coil_125", "Coil_126", "Coil_127", _
  "Coil_128", "Coil_129", "Coil_130", "Coil_131", "Coil_132", "Coil_133", _
  "Coil_134", "Coil_135", "Coil_136", "Coil_137", "Coil_138", "Coil_139", _
  "Coil_140", "Coil_141", "Coil_142", "Coil_143", "Coil_144", "Coil_145", _
  "Coil_146", "Coil_147", "Coil_148", "Coil_149", "Coil_150", "Coil_151", _
  "Coil_152", "Coil_153", "Coil_154", "Coil_155", "Coil_156", "Coil_157", _
  "Coil_158", "Coil_159", "Coil_160", "Coil_161", "Coil_162", "Coil_163", _
  "Coil_164", "Coil_165", "Coil_166", "Coil_167", "Coil_168", "Coil_169", _
  "Coil_170", "Coil_171", "Coil_172", "Coil_173", "Coil_174", "Coil_175", _
  "Coil_176", "Coil_177", "Coil_178", "Coil_179", "Coil_180", "Coil_181", _
  "Coil_182", "Coil_183", "Coil_184", "Coil_185", "Coil_186", "Coil_187", _
  "Coil_188", "Coil_189", "Coil_190", "Coil_191", "Coil_192", "Coil_193", _
  "Coil_194", "Coil_195", "Coil_196", "Coil_197", "Coil_198", "Coil_199", _
  "Coil_200", "Coil_201", "Coil_202", "Coil_203", "Coil_204", "Coil_205", _
  "Coil_206", "Coil_207", "Coil_208", "Coil_209", "Coil_210", "Coil_211", _
  "Coil_212", "Coil_213", "Coil_214", "Coil_215", "Coil_216", "Coil_217", _
  "Coil_218", "Coil_219", "Coil_220", "Coil_221", "Coil_222", "Coil_223", _
  "Coil_224", "Coil_225", "Coil_226", "Coil_227", "Coil_228", "Coil_229", _
  "Coil_230", "Coil_231", "Coil_232", "Coil_233", "Coil_234", "Coil_235", _
  "Coil_236", "Coil_237", "Coil_238", "Coil_239", "Coil_240", "Coil_241", _
  "Coil_242", "Coil_243", "Coil_244", "Coil_245", "Coil_246", "Coil_247", _
  "Coil_248", "Coil_249", "Coil_250", "Coil_251", "Coil_252", "Coil_253", _
  "Coil_254", "Coil_255", "Coil_256", "Coil_257", "Coil_258", "Coil_259", _
  "Coil_260", "Coil_261", "Coil_262", "Coil_263", "Coil_264", "Coil_265", _
  "Coil_266", "Coil_267", "Coil_268", "Coil_269", "Coil_270", "Coil_271", _
  "Coil_272", "Coil_273", "Coil_274", "Coil_275", "Coil_276", "Coil_277", _
  "Coil_278", "Coil_279", "Coil_280", "Coil_281", "Coil_282", "Coil_283", _
  "Coil_284", "Coil_285", "Coil_286", "Coil_287", "Coil_288", "Coil_289", _
  "Coil_290", "Coil_291", "Coil_292", "Coil_293", "Coil_294", "Coil_295", _
  "Coil_296", "Coil_297", "Coil_298", "Coil_299"), "NormalDevChoice:=", 2, _
  "NormalDev:=", "30deg", "AspectRatioChoice:=", 1)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "0"))), Array("NAME:Attributes", "Name:=", "Rotor", "Flags:=", _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "M22_24G_3DSF0.960", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "2"))), Array("NAME:Attributes", "Name:=", "Field", "Flags:=", _
  "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Field"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "11.25deg", _
  "NumClones:=", "32"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Field"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Field_0"))))
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "4"))), Array("NAME:Attributes", "Name:=", "FieldTerm", _
  "Flags:=", "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "FieldTerm:CreateUserDefinedPart:1", "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", _
  "FieldTerm"), Array("NAME:DuplicateAroundAxisParameters", _
  "CoordinateSystemID:=", -1, "CreateNewObjects:=", true, "WhichAxis:=", "Z", _
  "AngleStr:=", "11.25deg", "NumClones:=", "32"), Array("NAME:Options", _
  "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "FieldTerm"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "FieldTerm_0"))))
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:Field", "Type:=", "Current", _
  "IsSolid:=", false, "Current:=", "203.396A", "Resistance:=", "0ohm", _
  "Inductance:=", "0H", "ParallelBranchesNum:=", "1")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_0", "Objects:=", Array(_
  "FieldTerm_0"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_1", "Objects:=", Array(_
  "FieldTerm_1"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_2", "Objects:=", Array(_
  "FieldTerm_2"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_3", "Objects:=", Array(_
  "FieldTerm_3"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_4", "Objects:=", Array(_
  "FieldTerm_4"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_5", "Objects:=", Array(_
  "FieldTerm_5"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_6", "Objects:=", Array(_
  "FieldTerm_6"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_7", "Objects:=", Array(_
  "FieldTerm_7"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_8", "Objects:=", Array(_
  "FieldTerm_8"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_9", "Objects:=", Array(_
  "FieldTerm_9"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_10", "Objects:=", Array(_
  "FieldTerm_10"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_11", "Objects:=", Array(_
  "FieldTerm_11"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_12", "Objects:=", Array(_
  "FieldTerm_12"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_13", "Objects:=", Array(_
  "FieldTerm_13"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_14", "Objects:=", Array(_
  "FieldTerm_14"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_15", "Objects:=", Array(_
  "FieldTerm_15"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_16", "Objects:=", Array(_
  "FieldTerm_16"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_17", "Objects:=", Array(_
  "FieldTerm_17"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_18", "Objects:=", Array(_
  "FieldTerm_18"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_19", "Objects:=", Array(_
  "FieldTerm_19"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_20", "Objects:=", Array(_
  "FieldTerm_20"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_21", "Objects:=", Array(_
  "FieldTerm_21"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_22", "Objects:=", Array(_
  "FieldTerm_22"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_23", "Objects:=", Array(_
  "FieldTerm_23"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_24", "Objects:=", Array(_
  "FieldTerm_24"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_25", "Objects:=", Array(_
  "FieldTerm_25"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_26", "Objects:=", Array(_
  "FieldTerm_26"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_27", "Objects:=", Array(_
  "FieldTerm_27"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_28", "Objects:=", Array(_
  "FieldTerm_28"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_29", "Objects:=", Array(_
  "FieldTerm_29"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTerm_30", "Objects:=", Array(_
  "FieldTerm_30"), "Conductor number:=", 150, "Point out of terminal:=", _
  false, "Winding:=", "Field")
oModule.AssignCoilTerminal Array("NAME:FieldTermRe_31", "Objects:=", Array(_
  "FieldTerm_31"), "Conductor number:=", 150, "Point out of terminal:=", _
  true, "Winding:=", "Field")
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Field", "RefineInside:=", true, _
  "Objects:=", Array("Field_0", "Field_1", "Field_2", "Field_3", "Field_4", _
  "Field_5", "Field_6", "Field_7", "Field_8", "Field_9", "Field_10", _
  "Field_11", "Field_12", "Field_13", "Field_14", "Field_15", "Field_16", _
  "Field_17", "Field_18", "Field_19", "Field_20", "Field_21", "Field_22", _
  "Field_23", "Field_24", "Field_25", "Field_26", "Field_27", "Field_28", _
  "Field_29", "Field_30", "Field_31"), "RestrictElem:=", false, _
  "NumMaxElem:=", "1000", "RestrictLength:=", true, "MaxLength:=", "217.2mm")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "3"))), Array("NAME:Attributes", "Name:=", "Bar", "Flags:=", "", _
  "Color:=", "(119 198 206)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "copper_100C", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Bar:CreateUserDefinedPart:1", _
  "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Bar", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
oEditor.SeparateBody  Array("NAME:Selections", "Selections:=", "Bar")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetEddyEffect Array("NAME:Eddy Effect Setting", Array(_
  "NAME:EddyEffectVector", Array("NAME:Data", "Object Name:=", "Bar", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate1", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate2", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate3", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate4", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate5", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate6", _
  "Eddy Effect:=", true), Array("NAME:Data", "Object Name:=", "Bar_Separate7", _
  "Eddy Effect:=", true)))
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Bar", "RefineInside:=", true, _
  "Objects:=", Array("Bar", "Bar_Separate1", "Bar_Separate2", "Bar_Separate3", _
  "Bar_Separate4", "Bar_Separate5", "Bar_Separate6", "Bar_Separate7"), _
  "RestrictElem:=", false, "NumMaxElem:=", "1000", "RestrictLength:=", true, _
  "MaxLength:=", "20mm")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5398mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "1800mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1180mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "2"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "10mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "5mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "1deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "2deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "355mm"), Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", _
  "40mm"), Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "240mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "300mm"), Array(_
  "NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Offset", "Value:=", "10mm"), Array("NAME:Pair", "Name:=", _
  "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Off2_y", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", "Value:=", _
  "41.5mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", "1"), _
  Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "2423.705951218544mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "100"))), Array("NAME:Attributes", "Name:=", "InnerRegion", _
  "Flags:=", "", "Color:=", "(0 255 255)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "vacuum", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "LenRegion", "1180mm + 2*endRegion"
On Error Goto 0
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignLengthOp Array("NAME:Length_Main", "RefineInside:=", true, _
  "Objects:=", Array("Stator", "Rotor", "Band", "OuterRegion", "InnerRegion", _
  "Shaft"), "RestrictElem:=", false, "NumMaxElem:=", "1000", _
  "RestrictLength:=", true, "MaxLength:=", "230.4mm")
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Band", _
  "OuterRegion", "InnerRegion", "MasterSheet", "SlaveSheet"), Array(_
  "NAME:ChangedProps", Array("NAME:Transparent", "Value:=", 0.75))))
oEditor.Move Array("NAME:Selections", "Selections:=", "CoilTerm_0,CoilTerm_1" & _
  ",CoilTerm_2,CoilTerm_3,CoilTerm_4,CoilTerm_5,CoilTerm_6,CoilTerm_7" & _
  ",CoilTerm_8,CoilTerm_9,CoilTerm_10,CoilTerm_11,CoilTerm_12,CoilTerm_13" & _
  ",CoilTerm_14,CoilTerm_15,CoilTerm_16,CoilTerm_17,CoilTerm_18" & _
  ",CoilTerm_19,CoilTerm_20,CoilTerm_21,CoilTerm_22,CoilTerm_23" & _
  ",CoilTerm_24,CoilTerm_25,CoilTerm_26,CoilTerm_27,CoilTerm_28" & _
  ",CoilTerm_29,CoilTerm_30,CoilTerm_31,CoilTerm_32,CoilTerm_33" & _
  ",CoilTerm_34,CoilTerm_35,CoilTerm_36,CoilTerm_37,CoilTerm_38" & _
  ",CoilTerm_39,CoilTerm_40,CoilTerm_41,CoilTerm_42,CoilTerm_43" & _
  ",CoilTerm_44,CoilTerm_45,CoilTerm_46,CoilTerm_47,CoilTerm_48" & _
  ",CoilTerm_49,CoilTerm_50,CoilTerm_51,CoilTerm_52,CoilTerm_53" & _
  ",CoilTerm_54,CoilTerm_55,CoilTerm_56,CoilTerm_57,CoilTerm_58" & _
  ",CoilTerm_59,CoilTerm_60,CoilTerm_61,CoilTerm_62,CoilTerm_63" & _
  ",CoilTerm_64,CoilTerm_65,CoilTerm_66,CoilTerm_67,CoilTerm_68" & _
  ",CoilTerm_69,CoilTerm_70,CoilTerm_71,CoilTerm_72,CoilTerm_73" & _
  ",CoilTerm_74,CoilTerm_75,CoilTerm_76,CoilTerm_77,CoilTerm_78" & _
  ",CoilTerm_79,CoilTerm_80,CoilTerm_81,CoilTerm_82,CoilTerm_83" & _
  ",CoilTerm_84,CoilTerm_85,CoilTerm_86,CoilTerm_87,CoilTerm_88" & _
  ",CoilTerm_89,CoilTerm_90,CoilTerm_91,CoilTerm_92,CoilTerm_93" & _
  ",CoilTerm_94,CoilTerm_95,CoilTerm_96,CoilTerm_97,CoilTerm_98" & _
  ",CoilTerm_99,CoilTerm_100,CoilTerm_101,CoilTerm_102,CoilTerm_103" & _
  ",CoilTerm_104,CoilTerm_105,CoilTerm_106,CoilTerm_107,CoilTerm_108" & _
  ",CoilTerm_109,CoilTerm_110,CoilTerm_111,CoilTerm_112,CoilTerm_113" & _
  ",CoilTerm_114,CoilTerm_115,CoilTerm_116,CoilTerm_117,CoilTerm_118" & _
  ",CoilTerm_119,CoilTerm_120,CoilTerm_121,CoilTerm_122,CoilTerm_123" & _
  ",CoilTerm_124,CoilTerm_125,CoilTerm_126,CoilTerm_127,CoilTerm_128" & _
  ",CoilTerm_129,CoilTerm_130,CoilTerm_131,CoilTerm_132,CoilTerm_133" & _
  ",CoilTerm_134,CoilTerm_135,CoilTerm_136,CoilTerm_137,CoilTerm_138" & _
  ",CoilTerm_139,CoilTerm_140,CoilTerm_141,CoilTerm_142,CoilTerm_143" & _
  ",CoilTerm_144,CoilTerm_145,CoilTerm_146,CoilTerm_147,CoilTerm_148" & _
  ",CoilTerm_149,CoilTerm_150,CoilTerm_151,CoilTerm_152,CoilTerm_153" & _
  ",CoilTerm_154,CoilTerm_155,CoilTerm_156,CoilTerm_157,CoilTerm_158" & _
  ",CoilTerm_159,CoilTerm_160,CoilTerm_161,CoilTerm_162,CoilTerm_163" & _
  ",CoilTerm_164,CoilTerm_165,CoilTerm_166,CoilTerm_167,CoilTerm_168" & _
  ",CoilTerm_169,CoilTerm_170,CoilTerm_171,CoilTerm_172,CoilTerm_173" & _
  ",CoilTerm_174,CoilTerm_175,CoilTerm_176,CoilTerm_177,CoilTerm_178" & _
  ",CoilTerm_179,CoilTerm_180,CoilTerm_181,CoilTerm_182,CoilTerm_183" & _
  ",CoilTerm_184,CoilTerm_185,CoilTerm_186,CoilTerm_187,CoilTerm_188" & _
  ",CoilTerm_189,CoilTerm_190,CoilTerm_191,CoilTerm_192,CoilTerm_193" & _
  ",CoilTerm_194,CoilTerm_195,CoilTerm_196,CoilTerm_197,CoilTerm_198" & _
  ",CoilTerm_199,CoilTerm_200,CoilTerm_201,CoilTerm_202,CoilTerm_203" & _
  ",CoilTerm_204,CoilTerm_205,CoilTerm_206,CoilTerm_207,CoilTerm_208" & _
  ",CoilTerm_209,CoilTerm_210,CoilTerm_211,CoilTerm_212,CoilTerm_213" & _
  ",CoilTerm_214,CoilTerm_215,CoilTerm_216,CoilTerm_217,CoilTerm_218" & _
  ",CoilTerm_219,CoilTerm_220,CoilTerm_221,CoilTerm_222,CoilTerm_223" & _
  ",CoilTerm_224,CoilTerm_225,CoilTerm_226,CoilTerm_227,CoilTerm_228" & _
  ",CoilTerm_229,CoilTerm_230,CoilTerm_231,CoilTerm_232,CoilTerm_233" & _
  ",CoilTerm_234,CoilTerm_235,CoilTerm_236,CoilTerm_237,CoilTerm_238" & _
  ",CoilTerm_239,CoilTerm_240,CoilTerm_241,CoilTerm_242,CoilTerm_243" & _
  ",CoilTerm_244,CoilTerm_245,CoilTerm_246,CoilTerm_247,CoilTerm_248" & _
  ",CoilTerm_249,CoilTerm_250,CoilTerm_251,CoilTerm_252,CoilTerm_253" & _
  ",CoilTerm_254,CoilTerm_255,CoilTerm_256,CoilTerm_257,CoilTerm_258" & _
  ",CoilTerm_259,CoilTerm_260,CoilTerm_261,CoilTerm_262,CoilTerm_263" & _
  ",CoilTerm_264,CoilTerm_265,CoilTerm_266,CoilTerm_267,CoilTerm_268" & _
  ",CoilTerm_269,CoilTerm_270,CoilTerm_271,CoilTerm_272,CoilTerm_273" & _
  ",CoilTerm_274,CoilTerm_275,CoilTerm_276,CoilTerm_277,CoilTerm_278" & _
  ",CoilTerm_279,CoilTerm_280,CoilTerm_281,CoilTerm_282,CoilTerm_283" & _
  ",CoilTerm_284,CoilTerm_285,CoilTerm_286,CoilTerm_287,CoilTerm_288" & _
  ",CoilTerm_289,CoilTerm_290,CoilTerm_291,CoilTerm_292,CoilTerm_293" & _
  ",CoilTerm_294,CoilTerm_295,CoilTerm_296,CoilTerm_297,CoilTerm_298" & _
  ",CoilTerm_299,FieldTerm_0,FieldTerm_1,FieldTerm_2,FieldTerm_3" & _
  ",FieldTerm_4,FieldTerm_5,FieldTerm_6,FieldTerm_7,FieldTerm_8" & _
  ",FieldTerm_9,FieldTerm_10,FieldTerm_11,FieldTerm_12,FieldTerm_13" & _
  ",FieldTerm_14,FieldTerm_15,FieldTerm_16,FieldTerm_17,FieldTerm_18" & _
  ",FieldTerm_19,FieldTerm_20,FieldTerm_21,FieldTerm_22,FieldTerm_23" & _
  ",FieldTerm_24,FieldTerm_25,FieldTerm_26,FieldTerm_27,FieldTerm_28" & _
  ",FieldTerm_29,FieldTerm_30,FieldTerm_31"), Array(_
  "NAME:TranslateParameters", "CoordinateSystemID:=", -1, _
  "TranslateVectorX:=", "0mm", "TranslateVectorY:=", "0mm", _
  "TranslateVectorZ:=", "266.5mm")
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Band,InnerRegion" & _
  ",Shaft,Stator,Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24" & _
  ",Coil_25,Coil_26,Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32" & _
  ",Coil_33,Coil_34,Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40" & _
  ",Coil_41,Coil_42,Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48" & _
  ",Coil_49,Coil_50,Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56" & _
  ",Coil_57,Coil_58,Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64" & _
  ",Coil_65,Coil_66,Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,Coil_72" & _
  ",Coil_73,Coil_74,Coil_75,Coil_76,Coil_77,Coil_78,Coil_79,Coil_80" & _
  ",Coil_81,Coil_82,Coil_83,Coil_84,Coil_85,Coil_86,Coil_87,Coil_88" & _
  ",Coil_89,Coil_90,Coil_91,Coil_92,Coil_93,Coil_94,Coil_95,Coil_96" & _
  ",Coil_97,Coil_98,Coil_99,Coil_100,Coil_101,Coil_102,Coil_103,Coil_104" & _
  ",Coil_105,Coil_106,Coil_107,Coil_108,Coil_109,Coil_110,Coil_111" & _
  ",Coil_112,Coil_113,Coil_114,Coil_115,Coil_116,Coil_117,Coil_118" & _
  ",Coil_119,Coil_120,Coil_121,Coil_122,Coil_123,Coil_124,Coil_125" & _
  ",Coil_126,Coil_127,Coil_128,Coil_129,Coil_130,Coil_131,Coil_132" & _
  ",Coil_133,Coil_134,Coil_135,Coil_136,Coil_137,Coil_138,Coil_139" & _
  ",Coil_140,Coil_141,Coil_142,Coil_143,Coil_144,Coil_145,Coil_146" & _
  ",Coil_147,Coil_148,Coil_149,Coil_150,Coil_151,Coil_152,Coil_153" & _
  ",Coil_154,Coil_155,Coil_156,Coil_157,Coil_158,Coil_159,Coil_160" & _
  ",Coil_161,Coil_162,Coil_163,Coil_164,Coil_165,Coil_166,Coil_167" & _
  ",Coil_168,Coil_169,Coil_170,Coil_171,Coil_172,Coil_173,Coil_174" & _
  ",Coil_175,Coil_176,Coil_177,Coil_178,Coil_179,Coil_180,Coil_181" & _
  ",Coil_182,Coil_183,Coil_184,Coil_185,Coil_186,Coil_187,Coil_188" & _
  ",Coil_189,Coil_190,Coil_191,Coil_192,Coil_193,Coil_194,Coil_195" & _
  ",Coil_196,Coil_197,Coil_198,Coil_199,Coil_200,Coil_201,Coil_202" & _
  ",Coil_203,Coil_204,Coil_205,Coil_206,Coil_207,Coil_208,Coil_209" & _
  ",Coil_210,Coil_211,Coil_212,Coil_213,Coil_214,Coil_215,Coil_216" & _
  ",Coil_217,Coil_218,Coil_219,Coil_220,Coil_221,Coil_222,Coil_223" & _
  ",Coil_224,Coil_225,Coil_226,Coil_227,Coil_228,Coil_229,Coil_230" & _
  ",Coil_231,Coil_232,Coil_233,Coil_234,Coil_235,Coil_236,Coil_237" & _
  ",Coil_238,Coil_239,Coil_240,Coil_241,Coil_242,Coil_243,Coil_244" & _
  ",Coil_245,Coil_246,Coil_247,Coil_248,Coil_249,Coil_250,Coil_251" & _
  ",Coil_252,Coil_253,Coil_254,Coil_255,Coil_256,Coil_257,Coil_258" & _
  ",Coil_259,Coil_260,Coil_261,Coil_262,Coil_263,Coil_264,Coil_265" & _
  ",Coil_266,Coil_267,Coil_268,Coil_269,Coil_270,Coil_271,Coil_272" & _
  ",Coil_273,Coil_274,Coil_275,Coil_276,Coil_277,Coil_278,Coil_279" & _
  ",Coil_280,Coil_281,Coil_282,Coil_283,Coil_284,Coil_285,Coil_286" & _
  ",Coil_287,Coil_288,Coil_289,Coil_290,Coil_291,Coil_292,Coil_293" & _
  ",Coil_294,Coil_295,Coil_296,Coil_297,Coil_298,Coil_299,CoilTerm_0" & _
  ",CoilTerm_1,CoilTerm_2,CoilTerm_3,CoilTerm_4,CoilTerm_5,CoilTerm_6" & _
  ",CoilTerm_7,CoilTerm_8,CoilTerm_9,CoilTerm_10,CoilTerm_11,CoilTerm_12" & _
  ",CoilTerm_13,CoilTerm_14,CoilTerm_15,CoilTerm_16,CoilTerm_17" & _
  ",CoilTerm_18,CoilTerm_19,CoilTerm_20,CoilTerm_21,CoilTerm_22" & _
  ",CoilTerm_23,CoilTerm_24,CoilTerm_25,CoilTerm_26,CoilTerm_27" & _
  ",CoilTerm_28,CoilTerm_29,CoilTerm_30,CoilTerm_31,CoilTerm_32" & _
  ",CoilTerm_33,CoilTerm_34,CoilTerm_35,CoilTerm_36,CoilTerm_37" & _
  ",CoilTerm_38,CoilTerm_39,CoilTerm_40,CoilTerm_41,CoilTerm_42" & _
  ",CoilTerm_43,CoilTerm_44,CoilTerm_45,CoilTerm_46,CoilTerm_47" & _
  ",CoilTerm_48,CoilTerm_49,CoilTerm_50,CoilTerm_51,CoilTerm_52" & _
  ",CoilTerm_53,CoilTerm_54,CoilTerm_55,CoilTerm_56,CoilTerm_57" & _
  ",CoilTerm_58,CoilTerm_59,CoilTerm_60,CoilTerm_61,CoilTerm_62" & _
  ",CoilTerm_63,CoilTerm_64,CoilTerm_65,CoilTerm_66,CoilTerm_67" & _
  ",CoilTerm_68,CoilTerm_69,CoilTerm_70,CoilTerm_71,CoilTerm_72" & _
  ",CoilTerm_73,CoilTerm_74,CoilTerm_75,CoilTerm_76,CoilTerm_77" & _
  ",CoilTerm_78,CoilTerm_79,CoilTerm_80,CoilTerm_81,CoilTerm_82" & _
  ",CoilTerm_83,CoilTerm_84,CoilTerm_85,CoilTerm_86,CoilTerm_87" & _
  ",CoilTerm_88,CoilTerm_89,CoilTerm_90,CoilTerm_91,CoilTerm_92" & _
  ",CoilTerm_93,CoilTerm_94,CoilTerm_95,CoilTerm_96,CoilTerm_97" & _
  ",CoilTerm_98,CoilTerm_99,CoilTerm_100,CoilTerm_101,CoilTerm_102" & _
  ",CoilTerm_103,CoilTerm_104,CoilTerm_105,CoilTerm_106,CoilTerm_107" & _
  ",CoilTerm_108,CoilTerm_109,CoilTerm_110,CoilTerm_111,CoilTerm_112" & _
  ",CoilTerm_113,CoilTerm_114,CoilTerm_115,CoilTerm_116,CoilTerm_117" & _
  ",CoilTerm_118,CoilTerm_119,CoilTerm_120,CoilTerm_121,CoilTerm_122" & _
  ",CoilTerm_123,CoilTerm_124,CoilTerm_125,CoilTerm_126,CoilTerm_127" & _
  ",CoilTerm_128,CoilTerm_129,CoilTerm_130,CoilTerm_131,CoilTerm_132" & _
  ",CoilTerm_133,CoilTerm_134,CoilTerm_135,CoilTerm_136,CoilTerm_137" & _
  ",CoilTerm_138,CoilTerm_139,CoilTerm_140,CoilTerm_141,CoilTerm_142" & _
  ",CoilTerm_143,CoilTerm_144,CoilTerm_145,CoilTerm_146,CoilTerm_147" & _
  ",CoilTerm_148,CoilTerm_149,CoilTerm_150,CoilTerm_151,CoilTerm_152" & _
  ",CoilTerm_153,CoilTerm_154,CoilTerm_155,CoilTerm_156,CoilTerm_157" & _
  ",CoilTerm_158,CoilTerm_159,CoilTerm_160,CoilTerm_161,CoilTerm_162" & _
  ",CoilTerm_163,CoilTerm_164,CoilTerm_165,CoilTerm_166,CoilTerm_167" & _
  ",CoilTerm_168,CoilTerm_169,CoilTerm_170,CoilTerm_171,CoilTerm_172" & _
  ",CoilTerm_173,CoilTerm_174,CoilTerm_175,CoilTerm_176,CoilTerm_177" & _
  ",CoilTerm_178,CoilTerm_179,CoilTerm_180,CoilTerm_181,CoilTerm_182" & _
  ",CoilTerm_183,CoilTerm_184,CoilTerm_185,CoilTerm_186,CoilTerm_187" & _
  ",CoilTerm_188,CoilTerm_189,CoilTerm_190,CoilTerm_191,CoilTerm_192" & _
  ",CoilTerm_193,CoilTerm_194,CoilTerm_195,CoilTerm_196,CoilTerm_197" & _
  ",CoilTerm_198,CoilTerm_199,CoilTerm_200,CoilTerm_201,CoilTerm_202" & _
  ",CoilTerm_203,CoilTerm_204,CoilTerm_205,CoilTerm_206,CoilTerm_207" & _
  ",CoilTerm_208,CoilTerm_209,CoilTerm_210,CoilTerm_211,CoilTerm_212" & _
  ",CoilTerm_213,CoilTerm_214,CoilTerm_215,CoilTerm_216,CoilTerm_217" & _
  ",CoilTerm_218,CoilTerm_219,CoilTerm_220,CoilTerm_221,CoilTerm_222" & _
  ",CoilTerm_223,CoilTerm_224,CoilTerm_225,CoilTerm_226,CoilTerm_227" & _
  ",CoilTerm_228,CoilTerm_229,CoilTerm_230,CoilTerm_231,CoilTerm_232" & _
  ",CoilTerm_233,CoilTerm_234,CoilTerm_235,CoilTerm_236,CoilTerm_237" & _
  ",CoilTerm_238,CoilTerm_239,CoilTerm_240,CoilTerm_241,CoilTerm_242" & _
  ",CoilTerm_243,CoilTerm_244,CoilTerm_245,CoilTerm_246,CoilTerm_247" & _
  ",CoilTerm_248,CoilTerm_249,CoilTerm_250,CoilTerm_251,CoilTerm_252" & _
  ",CoilTerm_253,CoilTerm_254,CoilTerm_255,CoilTerm_256,CoilTerm_257" & _
  ",CoilTerm_258,CoilTerm_259,CoilTerm_260,CoilTerm_261,CoilTerm_262" & _
  ",CoilTerm_263,CoilTerm_264,CoilTerm_265,CoilTerm_266,CoilTerm_267" & _
  ",CoilTerm_268,CoilTerm_269,CoilTerm_270,CoilTerm_271,CoilTerm_272" & _
  ",CoilTerm_273,CoilTerm_274,CoilTerm_275,CoilTerm_276,CoilTerm_277" & _
  ",CoilTerm_278,CoilTerm_279,CoilTerm_280,CoilTerm_281,CoilTerm_282" & _
  ",CoilTerm_283,CoilTerm_284,CoilTerm_285,CoilTerm_286,CoilTerm_287" & _
  ",CoilTerm_288,CoilTerm_289,CoilTerm_290,CoilTerm_291,CoilTerm_292" & _
  ",CoilTerm_293,CoilTerm_294,CoilTerm_295,CoilTerm_296,CoilTerm_297" & _
  ",CoilTerm_298,CoilTerm_299,Rotor,Field_0,Field_1,Field_2,Field_3" & _
  ",Field_4,Field_5,Field_6,Field_7,Field_8,Field_9,Field_10,Field_11" & _
  ",Field_12,Field_13,Field_14,Field_15,Field_16,Field_17,Field_18" & _
  ",Field_19,Field_20,Field_21,Field_22,Field_23,Field_24,Field_25" & _
  ",Field_26,Field_27,Field_28,Field_29,Field_30,Field_31,FieldTerm_0" & _
  ",FieldTerm_1,FieldTerm_2,FieldTerm_3,FieldTerm_4,FieldTerm_5" & _
  ",FieldTerm_6,FieldTerm_7,FieldTerm_8,FieldTerm_9,FieldTerm_10" & _
  ",FieldTerm_11,FieldTerm_12,FieldTerm_13,FieldTerm_14,FieldTerm_15" & _
  ",FieldTerm_16,FieldTerm_17,FieldTerm_18,FieldTerm_19,FieldTerm_20" & _
  ",FieldTerm_21,FieldTerm_22,FieldTerm_23,FieldTerm_24,FieldTerm_25" & _
  ",FieldTerm_26,FieldTerm_27,FieldTerm_28,FieldTerm_29,FieldTerm_30" & _
  ",FieldTerm_31,Bar,Bar_Separate1,Bar_Separate2,Bar_Separate3" & _
  ",Bar_Separate4,Bar_Separate5,Bar_Separate6,Bar_Separate7", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.AssignBand Array("NAME:MotionSetup1", "Move Type:=", "Rotate", _
  "Coordinate System:=", "Global", "Axis:=", "Z", "Is Positive:=", true, _
  "InitPos:=", "13.425deg", "HasRotateLimit:=", false, "NonCylindrical:=", _
  false, "Consider Mechanical Transient:=", true, "Angular Velocity:=", _
  "187.5rpm", "Moment of Inertia:=", 715186, "Damping:=", 1465.24, _
  "Load Torque:=", _
  "if(speed<19.635, -102654*speed, -3.95763e+07/speed) - 102654*(speed-19.635)*10", _
  "Objects:=", Array("Band"))
oModule.EditMotionSetup Array("NAME:Data", "Consider Mechanical Transient:=", _
  false)
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "Transient", Array("NAME:Setup1", "StopTime:=", "0.2s", _
  "TimeStep:=", "0.0002s")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetMinimumTimeStep "0.002ms" 
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "Torque", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Moving1.Torque")), Array()
oModule.CreateReport "Currents", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Current(PhaseA)", "Current(PhaseB)", _
  "Current(PhaseC)")), Array()
oModule.CreateReport "Induced Voltages", "Transient", "XY Plot", _
  "Setup1 : Transient", Array(), Array("Time:=", Array("All")), Array(_
  "X Component:=", "Time", "Y Component:=", Array("InducedVoltage(PhaseA)", _
  "InducedVoltage(PhaseB)", "InducedVoltage(PhaseC)")), Array()
oModule.CreateReport "Flux Linkages", "Transient", "XY Plot", _
  "Setup1 : Transient", Array(), Array("Time:=", Array("All")), Array(_
  "X Component:=", "Time", "Y Component:=", Array("FluxLinkage(PhaseA)", _
  "FluxLinkage(PhaseB)", "FluxLinkage(PhaseC)")), Array()
set oUnclassified = oEditor.GetObjectsInGroup("Unclassified")
Dim oObject
For Each oObject in oUnclassified
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", oObject), Array(_
  "NAME:ChangedProps", Array("NAME:Model", "Value:=", false))))
Next
oEditor.ShowWindow
Set oModule = oDesign.GetModule("OutputVariable")
oModule.CreateOutputVariable "pos", _
  "(Moving1.Position -13.425 * PI/180) * 16 + PI", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos0", "cos(pos)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos1", "cos(pos-2*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos2", "cos(pos-4*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin0", "sin(pos)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin1", "sin(pos-2*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin2", "sin(pos-4*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "Lad", _
  "L(PhaseA,PhaseA)*cos0 + L(PhaseA,PhaseB)*cos1 + L(PhaseA,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Laq", _
  "L(PhaseA,PhaseA)*sin0 + L(PhaseA,PhaseB)*sin1 + L(PhaseA,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lbd", _
  "L(PhaseB,PhaseA)*cos0 + L(PhaseB,PhaseB)*cos1 + L(PhaseB,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lbq", _
  "L(PhaseB,PhaseA)*sin0 + L(PhaseB,PhaseB)*sin1 + L(PhaseB,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lcd", _
  "L(PhaseC,PhaseA)*cos0 + L(PhaseC,PhaseB)*cos1 + L(PhaseC,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lcq", _
  "L(PhaseC,PhaseA)*sin0 + L(PhaseC,PhaseB)*sin1 + L(PhaseC,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "L_d", "(Lad*cos0 + Lbd*cos1 + Lcd*cos2) * 2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "L_q", "(Laq*sin0 + Lbq*sin1 + Lcq*sin2) * 2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Flux_d", _
  "(FluxLinkage(PhaseA)*cos0+FluxLinkage(PhaseB)*cos1+FluxLinkage(PhaseC)*cos2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Flux_q", _
  "(FluxLinkage(PhaseA)*sin0+FluxLinkage(PhaseB)*sin1+FluxLinkage(PhaseC)*sin2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "I_d", _
  "(Current(PhaseA)*cos0 + Current(PhaseB)*cos1 + Current(PhaseC)*cos2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "I_q", _
  "(Current(PhaseA)*sin0 + Current(PhaseB)*sin1 + Current(PhaseC)*sin2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Irms", "sqrt(I_d^2+I_q^2)/sqrt(2)", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Pcu", "3*Irms^2*0", "Setup1 : Transient", _
  "Transient", Array() 
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "L_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("L_d", "L_q")), Array()
oModule.CreateReport "Flux_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Flux_d", "Flux_q")), Array()
oModule.CreateReport "I_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("I_d", "I_q")), Array()
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "ComputeTransientInductance:=", true, "ComputeIncrementalMatrix:=", false)
