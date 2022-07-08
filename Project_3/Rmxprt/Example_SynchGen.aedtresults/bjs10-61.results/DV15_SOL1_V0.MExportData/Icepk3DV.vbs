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
if (oArgs.Count > 0) then 
  Set oDesign = oProject.GetDesign(oArgs(0))
else
  Set oDesign = oProject.GetActiveDesign()
end if
designName = oDesign.GetName
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", "mm", _
  "Rescale:=", false)
oDesign.SetSolutionType "Transient"
Set oModule = oDesign.GetModule("BoundarySetup")
oEditor.SetModelValidationSettings Array("NAME:Validation Options", _
  "EntityCheckLevel:=", "Strict", "IgnoreUnclassifiedObjects:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "InsulatorThreshold:=", 2.5e+06)
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:fractions", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:halfAxial", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "0"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:endRegion", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "501.27732402191668mm"))))
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
if (oDefinitionManager.DoesMaterialExist("steel_1008_3DSF0.960")) then
else
oDefinitionManager.AddMaterial Array("NAME:steel_1008_3DSF0.960", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.2402), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.8654), Array("NAME:Coordinate", "X:=", 477.5, "Y:=", _
  1.1106), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.2458), Array(_
  "NAME:Coordinate", "X:=", 795.8, "Y:=", 1.331), Array("NAME:Coordinate", _
  "X:=", 1591.5, "Y:=", 1.5), Array("NAME:Coordinate", "X:=", 3183.1, "Y:=", _
  1.6), Array("NAME:Coordinate", "X:=", 4774.6, "Y:=", 1.683), Array(_
  "NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.741), Array("NAME:Coordinate", _
  "X:=", 7957.7, "Y:=", 1.78), Array("NAME:Coordinate", "X:=", 15915.5, "Y:=", _
  1.905), Array("NAME:Coordinate", "X:=", 31831, "Y:=", 2.025), Array(_
  "NAME:Coordinate", "X:=", 47746.5, "Y:=", 2.085), Array("NAME:Coordinate", _
  "X:=", 63662, "Y:=", 2.13), Array("NAME:Coordinate", "X:=", 79577.5, "Y:=", _
  2.165), Array("NAME:Coordinate", "X:=", 159155, "Y:=", 2.28), Array(_
  "NAME:Coordinate", "X:=", 318310, "Y:=", 2.485), Array("NAME:Coordinate", _
  "X:=", 397887, "Y:=", 2.5851), Array("NAME:Coordinate", "X:=", 1.19366e+06, _
  "Y:=", 3.5861))), "conductivity:=", 2e+06, "mass_density:=", 7872, Array(_
  "NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.96", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
end if
if (oDefinitionManager.DoesMaterialExist("steel_1008_3DSF0.960")) then
else
oDefinitionManager.AddMaterial Array("NAME:steel_1008_3DSF0.960", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.2402), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.8654), Array("NAME:Coordinate", "X:=", 477.5, "Y:=", _
  1.1106), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.2458), Array(_
  "NAME:Coordinate", "X:=", 795.8, "Y:=", 1.331), Array("NAME:Coordinate", _
  "X:=", 1591.5, "Y:=", 1.5), Array("NAME:Coordinate", "X:=", 3183.1, "Y:=", _
  1.6), Array("NAME:Coordinate", "X:=", 4774.6, "Y:=", 1.683), Array(_
  "NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.741), Array("NAME:Coordinate", _
  "X:=", 7957.7, "Y:=", 1.78), Array("NAME:Coordinate", "X:=", 15915.5, "Y:=", _
  1.905), Array("NAME:Coordinate", "X:=", 31831, "Y:=", 2.025), Array(_
  "NAME:Coordinate", "X:=", 47746.5, "Y:=", 2.085), Array("NAME:Coordinate", _
  "X:=", 63662, "Y:=", 2.13), Array("NAME:Coordinate", "X:=", 79577.5, "Y:=", _
  2.165), Array("NAME:Coordinate", "X:=", 159155, "Y:=", 2.28), Array(_
  "NAME:Coordinate", "X:=", 318310, "Y:=", 2.485), Array("NAME:Coordinate", _
  "X:=", 397887, "Y:=", 2.5851), Array("NAME:Coordinate", "X:=", 1.19366e+06, _
  "Y:=", 3.5861))), "conductivity:=", 2e+06, "mass_density:=", 7872, Array(_
  "NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.96", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
end if
if (oDefinitionManager.DoesMaterialExist("copper_115C")) then
else
oDefinitionManager.AddMaterial Array("NAME:copper_115C", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), "conductivity:=", "4.07429e+07")
end if
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "755.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "Shaft", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "Length", "840mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "755.5mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "OuterRegion", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Length", "840mm + 2*endRegion"
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
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "1"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "Tool"))))
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.SetSymmetryMultiplier "fractions*(1+halfAxial)"
Set oModule = oDesign.GetModule("BoundarySetup")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/VentSlotCore", "Version:=", "12.1", "NoOfParameters:=", _
  27, "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "760mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "72"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "3.2mm"), Array("NAME:Pair", _
  "Name:=", "Hs2", "Value:=", "67mm"), Array("NAME:Pair", "Name:=", "Bs0", _
  "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "14.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14.7mm"), Array(_
  "NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HalfSlot", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "VentHoles", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "HoleDiaIn", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "HoleDiaOut", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "HoleLocIn", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "HoleLocOut", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "VentDucts", _
  "Value:=", "13"), Array("NAME:Pair", "Name:=", "DuctWidth", "Value:=", _
  "9.5mm"), Array("NAME:Pair", "Name:=", "DuctPitch", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "30deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "1762.554648043833mm"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Stator", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "M22_24G_3DSF0.900", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/VentSlotCore", "Version:=", "12.1", "NoOfParameters:=", _
  27, "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "760mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "72"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs01", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "3.2mm"), Array("NAME:Pair", _
  "Name:=", "Hs2", "Value:=", "67mm"), Array("NAME:Pair", "Name:=", "Bs0", _
  "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "14.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "14.7mm"), Array(_
  "NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HalfSlot", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "VentHoles", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "HoleDiaIn", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "HoleDiaOut", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "HoleLocIn", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "HoleLocOut", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "VentDucts", _
  "Value:=", "13"), Array("NAME:Pair", "Name:=", "DuctWidth", "Value:=", _
  "9.5mm"), Array("NAME:Pair", "Name:=", "DuctPitch", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "30deg"), Array(_
  "NAME:Pair", "Name:=", "LenRegion", "Value:=", "1762.554648043833mm"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "1"))), Array(_
  "NAME:Attributes", "Name:=", "StatorSlotInsu", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "LenRegion", _
  "840mm + 2*endRegion"
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:delta", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "-28.514673505590345deg"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:conds", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "5"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:R1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.2027877483906807ohm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:Le1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.0023052266650856533H"))))
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "762mm"), Array("NAME:Pair", "Name:=", _
  "DiaYoke", "Value:=", "1102.2mm"), Array("NAME:Pair", "Name:=", "Length", _
  "Value:=", "760mm"), Array("NAME:Pair", "Name:=", "Skew", "Value:=", "0deg"), _
  Array("NAME:Pair", "Name:=", "Slots", "Value:=", "72"), Array("NAME:Pair", _
  "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", "Name:=", "Hs0", _
  "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "3.2mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "67mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "14.7mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "14.7mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", "Name:=", _
  "CoilPitch", "Value:=", "10"), Array("NAME:Pair", "Name:=", "EndExt", _
  "Value:=", "105mm"), Array("NAME:Pair", "Name:=", "SpanExt", "Value:=", _
  "0.1mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", "10deg"), _
  Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "1762.554648043833mm"), _
  Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", "1"))), Array(_
  "NAME:Attributes", "Name:=", "Coil", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Coil"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "5deg", _
  "NumClones:=", "72"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Coil"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Coil_0"))))
On Error Resume Next
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "StatorSlotInsu", _
  "Tool Parts:=", "Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24" & _
  ",Coil_25,Coil_26,Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32" & _
  ",Coil_33,Coil_34,Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40" & _
  ",Coil_41,Coil_42,Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48" & _
  ",Coil_49,Coil_50,Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56" & _
  ",Coil_57,Coil_58,Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64" & _
  ",Coil_65,Coil_66,Coil_67,Coil_68,Coil_69,Coil_70,Coil_71"), Array(_
  "NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  true)
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "StatorSlotInsu", _
  "Tool Parts:=", "Stator"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "0"))), Array("NAME:Attributes", "Name:=", "Rotor", "Flags:=", _
  "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", _
  "steel_1008_3DSF0.960", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "2"))), Array("NAME:Attributes", "Name:=", "Field", "Flags:=", _
  "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Field"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "60deg", _
  "NumClones:=", "6"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Field"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Field_0"))))
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "749mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "250mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "840mm"), Array("NAME:Pair", "Name:=", _
  "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", _
  "6"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "1"), Array(_
  "NAME:Pair", "Name:=", "Hs0", "Value:=", "0.8mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0.8mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "3mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "14mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "7mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "5.064deg"), Array(_
  "NAME:Pair", "Name:=", "CenterPitch", "Value:=", "14.736deg"), Array(_
  "NAME:Pair", "Name:=", "Poles", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "WidthShoe", "Value:=", "280mm"), Array("NAME:Pair", "Name:=", _
  "HeightShoe", "Value:=", "53mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "146mm"), Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", _
  "147mm"), Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "43mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "43mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "40mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "21mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "1762.554648043833mm"), Array("NAME:Pair", "Name:=", "InfoCore", _
  "Value:=", "3"))), Array("NAME:Attributes", "Name:=", "Bar", "Flags:=", "", _
  "Color:=", "(119 198 206)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "copper_115C", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Bar:CreateUserDefinedPart:1", _
  "LenRegion", "840mm + 2*endRegion"
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
  "Eddy Effect:=", true)))
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion"), _
  Array("NAME:ChangedProps", Array("NAME:Transparent", "Value:=", 0.75))))
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Shaft,Stator,Coil_0" & _
  ",Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7,Coil_8,Coil_9,Coil_10" & _
  ",Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16,Coil_17,Coil_18" & _
  ",Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24,Coil_25,Coil_26" & _
  ",Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32,Coil_33,Coil_34" & _
  ",Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40,Coil_41,Coil_42" & _
  ",Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48,Coil_49,Coil_50" & _
  ",Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56,Coil_57,Coil_58" & _
  ",Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64,Coil_65,Coil_66" & _
  ",Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,Rotor,Field_0,Field_1,Field_2" & _
  ",Field_3,Field_4,Field_5,Bar,Bar_Separate1,Bar_Separate2,Bar_Separate3" & _
  ",Bar_Separate4,Bar_Separate5", "Tool Parts:=", "Tool"), Array(_
  "NAME:SubtractParameters", "CoordinateSystemID:=", -1, "KeepOriginals:=", _
  false)
oEditor.FitAll
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "Transient", Array("NAME:Setup1", "StopTime:=", "0.2s", _
  "TimeStep:=", "0.0002s")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetMinimumTimeStep "0.002ms" 
set oUnclassified = oEditor.GetObjectsInGroup("Unclassified")
Dim oObject
For Each oObject in oUnclassified
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", oObject), Array(_
  "NAME:ChangedProps", Array("NAME:Model", "Value:=", false))))
Next
