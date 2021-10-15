VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CreatePart()

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Add("Part")

End Sub

Sub SketchConnectorBase()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim originElements1 As OriginElements
Set originElements1 = part1.OriginElements

Dim reference1 As Reference
Set reference1 = originElements1.PlaneYZ

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 0#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = 1#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("Repère")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("Axe horizontal")

line2D1.ReportName = 1

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("Axe vertical")

line2D2.ReportName = 2

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(50#, 0#)

point2D1.ReportName = 3

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(0#, 0#, 50#, 0#)

line2D3.ReportName = 4

Dim point2D2 As Point2D
Set point2D2 = axis2D1.GetItem("Origine")

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D1

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(50#, 50#)

point2D3.ReportName = 5

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(50#, 0#, 50#, 50#)

line2D4.ReportName = 6

line2D4.StartPoint = point2D1

line2D4.EndPoint = point2D3

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(0#, 50#)

point2D4.ReportName = 7

Dim line2D5 As Line2D
Set line2D5 = factory2D1.CreateLine(50#, 50#, 0#, 50#)

line2D5.ReportName = 8

line2D5.StartPoint = point2D3

line2D5.EndPoint = point2D4

Dim line2D6 As Line2D
Set line2D6 = factory2D1.CreateLine(0#, 50#, 0#, 0#)

line2D6.ReportName = 9

line2D6.StartPoint = point2D4

line2D6.EndPoint = point2D2

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(line2D3)

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(line2D1)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference2, reference3)

constraint1.Mode = catCstModeDrivingDimension

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(line2D5)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference4, reference5)

constraint2.Mode = catCstModeDrivingDimension

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromObject(line2D4)

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(line2D2)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddBiEltCst(catCstTypeVerticality, reference6, reference7)

constraint3.Mode = catCstModeDrivingDimension

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromObject(line2D6)

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(line2D2)

Dim constraint4 As Constraint
Set constraint4 = constraints1.AddBiEltCst(catCstTypeVerticality, reference8, reference9)

constraint4.Mode = catCstModeDrivingDimension

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromObject(line2D6)

Dim constraint5 As Constraint
Set constraint5 = constraints1.AddMonoEltCst(catCstTypeLength, reference10)

constraint5.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint5.Dimension

length1.Value = 50#

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromObject(line2D5)

Dim constraint6 As Constraint
Set constraint6 = constraints1.AddMonoEltCst(catCstTypeLength, reference11)

constraint6.Mode = catCstModeDrivingDimension

Dim length2 As Length
Set length2 = constraint6.Dimension

length2.Value = 50#

sketch1.CloseEdition

part1.InWorkObject = body1

part1.Update

End Sub


Sub ExtrudeConnectorBase()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

part1.InWorkObject = body1

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Esquisse.1")

Dim pad1 As Pad
Set pad1 = shapeFactory1.AddNewPad(sketch1, 20#)

Dim limit1 As Limit
Set limit1 = pad1.FirstLimit

Dim length1 As Length
Set length1 = limit1.Dimension

length1.Value = 76#

part1.Update

End Sub

Sub SketchConnectorPlug()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(Pad.1;2);None:();Cf11:());Pad.1_ResultOUT;Z0;G4711)")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 76#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = 1#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("Repère")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("Axe horizontal")

line2D1.ReportName = 10

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("Axe vertical")

line2D2.ReportName = 11

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(25#, 25#)

point2D1.ReportName = 12

Dim point2D2 As Point2D
Set point2D2 = factory2D1.CreatePoint(45#, 45#)

point2D2.ReportName = 13

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(45#, 5#)

point2D3.ReportName = 14

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(45#, 45#, 45#, 5#)

line2D3.ReportName = 15

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D3

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(5#, 5#)

point2D4.ReportName = 16

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(45#, 5#, 5#, 5#)

line2D4.ReportName = 17

line2D4.StartPoint = point2D3

line2D4.EndPoint = point2D4

Dim point2D5 As Point2D
Set point2D5 = factory2D1.CreatePoint(5#, 45#)

point2D5.ReportName = 18

Dim line2D5 As Line2D
Set line2D5 = factory2D1.CreateLine(5#, 5#, 5#, 45#)

line2D5.ReportName = 19

line2D5.StartPoint = point2D4

line2D5.EndPoint = point2D5

Dim line2D6 As Line2D
Set line2D6 = factory2D1.CreateLine(5#, 45#, 45#, 45#)

line2D6.ReportName = 20

line2D6.StartPoint = point2D5

line2D6.EndPoint = point2D2

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(line2D3)

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(line2D2)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeVerticality, reference2, reference3)

constraint1.Mode = catCstModeDrivingDimension

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(line2D4)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference4, reference5)

constraint2.Mode = catCstModeDrivingDimension

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromObject(line2D5)

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(line2D2)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddBiEltCst(catCstTypeVerticality, reference6, reference7)

constraint3.Mode = catCstModeDrivingDimension

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromObject(line2D6)

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(line2D1)

Dim constraint4 As Constraint
Set constraint4 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference8, reference9)

constraint4.Mode = catCstModeDrivingDimension

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromObject(line2D3)

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromObject(line2D5)

Dim reference12 As Reference
Set reference12 = part1.CreateReferenceFromObject(point2D1)

Dim constraint5 As Constraint
Set constraint5 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference10, reference11, reference12)

constraint5.Mode = catCstModeDrivingDimension

Dim reference13 As Reference
Set reference13 = part1.CreateReferenceFromObject(line2D4)

Dim reference14 As Reference
Set reference14 = part1.CreateReferenceFromObject(line2D6)

Dim reference15 As Reference
Set reference15 = part1.CreateReferenceFromObject(point2D1)

Dim constraint6 As Constraint
Set constraint6 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference13, reference14, reference15)

constraint6.Mode = catCstModeDrivingDimension

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim pad1 As Pad
Set pad1 = shapes1.Item("Extrusion.1")

Dim reference16 As Reference
Set reference16 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;4)));None:();Cf11:());Face:(Brp:(Pad.1;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateProjections(reference16)

Dim geometry2D1 As Geometry2D
Set geometry2D1 = geometricElements2.Item("Empreinte.1")

geometry2D1.Construction = True

Dim reference17 As Reference
Set reference17 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;8)));None:();Cf11:());Face:(Brp:(Pad.1;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements3 As GeometricElements
Set geometricElements3 = factory2D1.CreateProjections(reference17)

Dim geometry2D2 As Geometry2D
Set geometry2D2 = geometricElements3.Item("Empreinte.1")

geometry2D2.Construction = True

Dim reference18 As Reference
Set reference18 = part1.CreateReferenceFromObject(geometry2D1)

Dim reference19 As Reference
Set reference19 = part1.CreateReferenceFromObject(geometry2D2)

Dim reference20 As Reference
Set reference20 = part1.CreateReferenceFromObject(point2D1)

Dim constraint7 As Constraint
Set constraint7 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference18, reference19, reference20)

constraint7.Mode = catCstModeDrivingDimension

Dim reference21 As Reference
Set reference21 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;9)));None:();Cf11:());Face:(Brp:(Pad.1;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements4 As GeometricElements
Set geometricElements4 = factory2D1.CreateProjections(reference21)

Dim geometry2D3 As Geometry2D
Set geometry2D3 = geometricElements4.Item("Empreinte.1")

geometry2D3.Construction = True

Dim reference22 As Reference
Set reference22 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;6)));None:();Cf11:());Face:(Brp:(Pad.1;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements5 As GeometricElements
Set geometricElements5 = factory2D1.CreateProjections(reference22)

Dim geometry2D4 As Geometry2D
Set geometry2D4 = geometricElements5.Item("Empreinte.1")

geometry2D4.Construction = True

Dim reference23 As Reference
Set reference23 = part1.CreateReferenceFromObject(geometry2D3)

Dim reference24 As Reference
Set reference24 = part1.CreateReferenceFromObject(geometry2D4)

Dim reference25 As Reference
Set reference25 = part1.CreateReferenceFromObject(point2D1)

Dim constraint8 As Constraint
Set constraint8 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference23, reference24, reference25)

constraint8.Mode = catCstModeDrivingDimension

Dim reference26 As Reference
Set reference26 = part1.CreateReferenceFromObject(line2D6)

Dim constraint9 As Constraint
Set constraint9 = constraints1.AddMonoEltCst(catCstTypeLength, reference26)

constraint9.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint9.Dimension

length1.Value = 40#

Dim reference27 As Reference
Set reference27 = part1.CreateReferenceFromObject(line2D3)

Dim constraint10 As Constraint
Set constraint10 = constraints1.AddMonoEltCst(catCstTypeLength, reference27)

constraint10.Mode = catCstModeDrivingDimension

Dim length2 As Length
Set length2 = constraint10.Dimension

length2.Value = 40#

Dim point2D6 As Point2D
Set point2D6 = factory2D1.CreatePoint(25#, 45#)

point2D6.ReportName = 21

Dim point2D7 As Point2D
Set point2D7 = factory2D1.CreatePoint(25#, 5#)

point2D7.ReportName = 22

Dim line2D7 As Line2D
Set line2D7 = factory2D1.CreateLine(25#, 45#, 25#, 5#)

line2D7.ReportName = 23

line2D7.Construction = True

line2D7.StartPoint = point2D6

line2D7.EndPoint = point2D7

Dim reference28 As Reference
Set reference28 = part1.CreateReferenceFromObject(point2D6)

Dim reference29 As Reference
Set reference29 = part1.CreateReferenceFromObject(line2D6)

Dim constraint11 As Constraint
Set constraint11 = constraints1.AddBiEltCst(catCstTypeMidPoint, reference28, reference29)

constraint11.Mode = catCstModeDrivingDimension

Dim reference30 As Reference
Set reference30 = part1.CreateReferenceFromObject(point2D7)

Dim reference31 As Reference
Set reference31 = part1.CreateReferenceFromObject(line2D4)

Dim constraint12 As Constraint
Set constraint12 = constraints1.AddBiEltCst(catCstTypeMidPoint, reference30, reference31)

constraint12.Mode = catCstModeDrivingDimension

Dim reference32 As Reference
Set reference32 = part1.CreateReferenceFromObject(line2D7)

Dim reference33 As Reference
Set reference33 = part1.CreateReferenceFromObject(line2D2)

Dim constraint13 As Constraint
Set constraint13 = constraints1.AddBiEltCst(catCstTypeVerticality, reference32, reference33)

constraint13.Mode = catCstModeDrivingDimension

Dim point2D8 As Point2D
Set point2D8 = factory2D1.CreatePoint(25#, 40#)

point2D8.ReportName = 24

Dim reference34 As Reference
Set reference34 = part1.CreateReferenceFromObject(point2D8)

Dim reference35 As Reference
Set reference35 = part1.CreateReferenceFromObject(line2D7)

Dim constraint14 As Constraint
Set constraint14 = constraints1.AddBiEltCst(catCstTypeOn, reference34, reference35)

constraint14.Mode = catCstModeDrivingDimension

Dim point2D9 As Point2D
Set point2D9 = factory2D1.CreatePoint(40#, 40#)

point2D9.ReportName = 25

Dim line2D8 As Line2D
Set line2D8 = factory2D1.CreateLine(25#, 40#, 40#, 40#)

line2D8.ReportName = 26

line2D8.StartPoint = point2D8

line2D8.EndPoint = point2D9

Dim reference36 As Reference
Set reference36 = part1.CreateReferenceFromObject(line2D8)

Dim reference37 As Reference
Set reference37 = part1.CreateReferenceFromObject(line2D1)

Dim constraint15 As Constraint
Set constraint15 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference36, reference37)

constraint15.Mode = catCstModeDrivingDimension

Dim point2D10 As Point2D
Set point2D10 = factory2D1.CreatePoint(40#, 13#)

point2D10.ReportName = 27

Dim line2D9 As Line2D
Set line2D9 = factory2D1.CreateLine(40#, 40#, 40#, 13#)

line2D9.ReportName = 28

line2D9.StartPoint = point2D9

line2D9.EndPoint = point2D10

Dim reference38 As Reference
Set reference38 = part1.CreateReferenceFromObject(line2D9)

Dim reference39 As Reference
Set reference39 = part1.CreateReferenceFromObject(line2D2)

Dim constraint16 As Constraint
Set constraint16 = constraints1.AddBiEltCst(catCstTypeVerticality, reference38, reference39)

constraint16.Mode = catCstModeDrivingDimension

Dim point2D11 As Point2D
Set point2D11 = factory2D1.CreatePoint(37#, 10#)

point2D11.ReportName = 29

Dim line2D10 As Line2D
Set line2D10 = factory2D1.CreateLine(40#, 13#, 37#, 10#)

line2D10.ReportName = 30

line2D10.StartPoint = point2D10

line2D10.EndPoint = point2D11

Dim point2D12 As Point2D
Set point2D12 = factory2D1.CreatePoint(25#, 10#)

point2D12.ReportName = 31

Dim line2D11 As Line2D
Set line2D11 = factory2D1.CreateLine(37#, 10#, 25#, 10#)

line2D11.ReportName = 32

line2D11.StartPoint = point2D11

line2D11.EndPoint = point2D12

Dim reference40 As Reference
Set reference40 = part1.CreateReferenceFromObject(point2D12)

Dim reference41 As Reference
Set reference41 = part1.CreateReferenceFromObject(line2D7)

Dim constraint17 As Constraint
Set constraint17 = constraints1.AddBiEltCst(catCstTypeOn, reference40, reference41)

constraint17.Mode = catCstModeDrivingDimension

Dim reference42 As Reference
Set reference42 = part1.CreateReferenceFromObject(line2D11)

Dim reference43 As Reference
Set reference43 = part1.CreateReferenceFromObject(line2D1)

Dim constraint18 As Constraint
Set constraint18 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference42, reference43)

constraint18.Mode = catCstModeDrivingDimension

Dim reference44 As Reference
Set reference44 = part1.CreateReferenceFromObject(line2D11)

Dim reference45 As Reference
Set reference45 = part1.CreateReferenceFromObject(line2D4)

Dim constraint19 As Constraint
Set constraint19 = constraints1.AddBiEltCst(catCstTypeDistance, reference44, reference45)

constraint19.Mode = catCstModeDrivingDimension

Dim length3 As Length
Set length3 = constraint19.Dimension

length3.Value = 5#

Dim reference46 As Reference
Set reference46 = part1.CreateReferenceFromObject(line2D8)

Dim reference47 As Reference
Set reference47 = part1.CreateReferenceFromObject(line2D6)

Dim constraint20 As Constraint
Set constraint20 = constraints1.AddBiEltCst(catCstTypeDistance, reference46, reference47)

constraint20.Mode = catCstModeDrivingDimension

Dim length4 As Length
Set length4 = constraint20.Dimension

length4.Value = 5#

Dim reference48 As Reference
Set reference48 = part1.CreateReferenceFromObject(line2D9)

Dim reference49 As Reference
Set reference49 = part1.CreateReferenceFromObject(line2D3)

Dim constraint21 As Constraint
Set constraint21 = constraints1.AddBiEltCst(catCstTypeDistance, reference48, reference49)

constraint21.Mode = catCstModeDrivingDimension

Dim length5 As Length
Set length5 = constraint21.Dimension

length5.Value = 5#

Dim reference50 As Reference
Set reference50 = part1.CreateReferenceFromObject(line2D10)

Dim reference51 As Reference
Set reference51 = part1.CreateReferenceFromObject(line2D11)

Dim constraint22 As Constraint
Set constraint22 = constraints1.AddBiEltCst(catCstTypeAngle, reference50, reference51)

constraint22.Mode = catCstModeDrivingDimension

constraint22.AngleSector = catCstAngleSector3

Dim angle1 As Angle
Set angle1 = constraint22.Dimension

angle1.Value = 45#

Dim reference52 As Reference
Set reference52 = part1.CreateReferenceFromObject(point2D11)

Dim reference53 As Reference
Set reference53 = part1.CreateReferenceFromObject(line2D9)

Dim constraint23 As Constraint
Set constraint23 = constraints1.AddBiEltCst(catCstTypeDistance, reference52, reference53)

constraint23.Mode = catCstModeDrivingDimension

Dim length6 As Length
Set length6 = constraint23.Dimension

length6.Value = 3#

Dim point2D13 As Point2D
Set point2D13 = factory2D1.CreatePoint(10#, 40#)

point2D13.ReportName = 1

Dim point2D14 As Point2D
Set point2D14 = factory2D1.CreatePoint(10#, 13#)

point2D14.ReportName = 2

Dim point2D15 As Point2D
Set point2D15 = factory2D1.CreatePoint(13#, 10#)

point2D15.ReportName = 3

Dim line2D12 As Line2D
Set line2D12 = factory2D1.CreateLine(10#, 40#, 10#, 13#)

line2D12.ReportName = 4

Dim line2D13 As Line2D
Set line2D13 = factory2D1.CreateLine(10#, 13#, 13#, 10#)

line2D13.ReportName = 5

Dim point2D16 As Point2D
Set point2D16 = factory2D1.CreatePoint(25#, 40#)

point2D16.ReportName = 6

Dim line2D14 As Line2D
Set line2D14 = factory2D1.CreateLine(25#, 40#, 10#, 40#)

line2D14.ReportName = 7

Dim point2D17 As Point2D
Set point2D17 = factory2D1.CreatePoint(25#, 10#)

point2D17.ReportName = 8

Dim line2D15 As Line2D
Set line2D15 = factory2D1.CreateLine(13#, 10#, 25#, 10#)

line2D15.ReportName = 9

Dim reference54 As Reference
Set reference54 = part1.CreateReferenceFromObject(point2D9)

Dim reference55 As Reference
Set reference55 = part1.CreateReferenceFromObject(point2D13)

Dim reference56 As Reference
Set reference56 = part1.CreateReferenceFromObject(line2D7)

Dim constraint24 As Constraint
Set constraint24 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference54, reference55, reference56)

constraint24.Mode = catCstModeDrivingDimension

Dim reference57 As Reference
Set reference57 = part1.CreateReferenceFromObject(point2D10)

Dim reference58 As Reference
Set reference58 = part1.CreateReferenceFromObject(point2D14)

Dim reference59 As Reference
Set reference59 = part1.CreateReferenceFromObject(line2D7)

Dim constraint25 As Constraint
Set constraint25 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference57, reference58, reference59)

constraint25.Mode = catCstModeDrivingDimension

Dim reference60 As Reference
Set reference60 = part1.CreateReferenceFromObject(point2D11)

Dim reference61 As Reference
Set reference61 = part1.CreateReferenceFromObject(point2D15)

Dim reference62 As Reference
Set reference62 = part1.CreateReferenceFromObject(line2D7)

Dim constraint26 As Constraint
Set constraint26 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference60, reference61, reference62)

constraint26.Mode = catCstModeDrivingDimension

Dim reference63 As Reference
Set reference63 = part1.CreateReferenceFromObject(line2D9)

Dim reference64 As Reference
Set reference64 = part1.CreateReferenceFromObject(line2D12)

Dim reference65 As Reference
Set reference65 = part1.CreateReferenceFromObject(line2D7)

Dim constraint27 As Constraint
Set constraint27 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference63, reference64, reference65)

constraint27.Mode = catCstModeDrivingDimension

Dim reference66 As Reference
Set reference66 = part1.CreateReferenceFromObject(line2D10)

Dim reference67 As Reference
Set reference67 = part1.CreateReferenceFromObject(line2D13)

Dim reference68 As Reference
Set reference68 = part1.CreateReferenceFromObject(line2D7)

Dim constraint28 As Constraint
Set constraint28 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference66, reference67, reference68)

constraint28.Mode = catCstModeDrivingDimension

Dim reference69 As Reference
Set reference69 = part1.CreateReferenceFromObject(point2D8)

Dim reference70 As Reference
Set reference70 = part1.CreateReferenceFromObject(point2D16)

Dim reference71 As Reference
Set reference71 = part1.CreateReferenceFromObject(line2D7)

Dim constraint29 As Constraint
Set constraint29 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference69, reference70, reference71)

constraint29.Mode = catCstModeDrivingDimension

Dim reference72 As Reference
Set reference72 = part1.CreateReferenceFromObject(line2D8)

Dim reference73 As Reference
Set reference73 = part1.CreateReferenceFromObject(line2D14)

Dim reference74 As Reference
Set reference74 = part1.CreateReferenceFromObject(line2D7)

Dim constraint30 As Constraint
Set constraint30 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference72, reference73, reference74)

constraint30.Mode = catCstModeDrivingDimension

Dim reference75 As Reference
Set reference75 = part1.CreateReferenceFromObject(point2D12)

Dim reference76 As Reference
Set reference76 = part1.CreateReferenceFromObject(point2D17)

Dim reference77 As Reference
Set reference77 = part1.CreateReferenceFromObject(line2D7)

Dim constraint31 As Constraint
Set constraint31 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference75, reference76, reference77)

constraint31.Mode = catCstModeDrivingDimension

Dim reference78 As Reference
Set reference78 = part1.CreateReferenceFromObject(line2D11)

Dim reference79 As Reference
Set reference79 = part1.CreateReferenceFromObject(line2D15)

Dim reference80 As Reference
Set reference80 = part1.CreateReferenceFromObject(line2D7)

Dim constraint32 As Constraint
Set constraint32 = constraints1.AddTriEltCst(catCstTypeSymmetry, reference78, reference79, reference80)

constraint32.Mode = catCstModeDrivingDimension

sketch1.CloseEdition

part1.InWorkObject = pad1

part1.Update

length3.Value = 5#

length4.Value = 5#

length5.Value = 5#

length6.Value = 3#

length3.Value = 5#

length4.Value = 5#

length5.Value = 5#

length6.Value = 3#

End Sub


Sub ExtrudeConnectorPlug()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Esquisse.2")

Dim pad1 As Pad
Set pad1 = shapeFactory1.AddNewPad(sketch1, 76#)

Dim limit1 As Limit
Set limit1 = pad1.FirstLimit

Dim length1 As Length
Set length1 = limit1.Dimension

length1.Value = 50#

part1.Update

End Sub

Sub SkecthConnectorEnd()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(Pad.1;1);None:();Cf11:());Pad.2_ResultOUT;Z0;G4711)")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 0#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = -1#
arrayOfVariantOfDouble1(5) = -0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("Repère")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("Axe horizontal")

line2D1.ReportName = 1

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("Axe vertical")

line2D2.ReportName = 2

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(-25#, 25#)

point2D1.ReportName = 3

Dim point2D2 As Point2D
Set point2D2 = factory2D1.CreatePoint(0#, 52.4)

point2D2.ReportName = 4

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(0#, -2.4)

point2D3.ReportName = 5

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(0#, 52.4, 0#, -2.4)

line2D3.ReportName = 6

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D3

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(-50#, -2.4)

point2D4.ReportName = 7

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(0#, -2.4, -50#, -2.4)

line2D4.ReportName = 8

line2D4.StartPoint = point2D3

line2D4.EndPoint = point2D4

Dim point2D5 As Point2D
Set point2D5 = factory2D1.CreatePoint(-50#, 52.4)

point2D5.ReportName = 9

Dim line2D5 As Line2D
Set line2D5 = factory2D1.CreateLine(-50#, -2.4, -50#, 52.4)

line2D5.ReportName = 10

line2D5.StartPoint = point2D4

line2D5.EndPoint = point2D5

Dim line2D6 As Line2D
Set line2D6 = factory2D1.CreateLine(-50#, 52.4, 0#, 52.4)

line2D6.ReportName = 11

line2D6.StartPoint = point2D5

line2D6.EndPoint = point2D2

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(line2D3)

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(line2D2)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeVerticality, reference2, reference3)

constraint1.Mode = catCstModeDrivingDimension

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(line2D4)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference4, reference5)

constraint2.Mode = catCstModeDrivingDimension

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromObject(line2D5)

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(line2D2)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddBiEltCst(catCstTypeVerticality, reference6, reference7)

constraint3.Mode = catCstModeDrivingDimension

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromObject(line2D6)

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(line2D1)

Dim constraint4 As Constraint
Set constraint4 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference8, reference9)

constraint4.Mode = catCstModeDrivingDimension

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromObject(line2D3)

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromObject(line2D5)

Dim reference12 As Reference
Set reference12 = part1.CreateReferenceFromObject(point2D1)

Dim constraint5 As Constraint
Set constraint5 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference10, reference11, reference12)

constraint5.Mode = catCstModeDrivingDimension

Dim reference13 As Reference
Set reference13 = part1.CreateReferenceFromObject(line2D4)

Dim reference14 As Reference
Set reference14 = part1.CreateReferenceFromObject(line2D6)

Dim reference15 As Reference
Set reference15 = part1.CreateReferenceFromObject(point2D1)

Dim constraint6 As Constraint
Set constraint6 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference13, reference14, reference15)

constraint6.Mode = catCstModeDrivingDimension

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim pad1 As Pad
Set pad1 = shapes1.Item("Extrusion.1")

Dim reference16 As Reference
Set reference16 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;1);None:();Cf11:());Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;6)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateProjections(reference16)

Dim geometry2D1 As Geometry2D
Set geometry2D1 = geometricElements2.Item("Empreinte.1")

geometry2D1.Construction = True

Dim reference17 As Reference
Set reference17 = part1.CreateReferenceFromObject(line2D5)

Dim reference18 As Reference
Set reference18 = part1.CreateReferenceFromObject(geometry2D1)

Dim constraint7 As Constraint
Set constraint7 = constraints1.AddBiEltCst(catCstTypeOn, reference17, reference18)

constraint7.Mode = catCstModeDrivingDimension

Dim reference19 As Reference
Set reference19 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;1);None:();Cf11:());Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;9)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements3 As GeometricElements
Set geometricElements3 = factory2D1.CreateProjections(reference19)

Dim geometry2D2 As Geometry2D
Set geometry2D2 = geometricElements3.Item("Empreinte.1")

geometry2D2.Construction = True

Dim reference20 As Reference
Set reference20 = part1.CreateReferenceFromObject(geometry2D2)

Dim reference21 As Reference
Set reference21 = part1.CreateReferenceFromObject(line2D3)

Dim constraint8 As Constraint
Set constraint8 = constraints1.AddBiEltCst(catCstTypeOn, reference20, reference21)

constraint8.Mode = catCstModeDrivingDimension

Dim reference22 As Reference
Set reference22 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;1);None:();Cf11:());Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;8)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements4 As GeometricElements
Set geometricElements4 = factory2D1.CreateProjections(reference22)

Dim geometry2D3 As Geometry2D
Set geometry2D3 = geometricElements4.Item("Empreinte.1")

geometry2D3.Construction = True

Dim reference23 As Reference
Set reference23 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;1);None:();Cf11:());Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;4)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements5 As GeometricElements
Set geometricElements5 = factory2D1.CreateProjections(reference23)

Dim geometry2D4 As Geometry2D
Set geometry2D4 = geometricElements5.Item("Empreinte.1")

geometry2D4.Construction = True

Dim reference24 As Reference
Set reference24 = part1.CreateReferenceFromObject(geometry2D3)

Dim reference25 As Reference
Set reference25 = part1.CreateReferenceFromObject(geometry2D4)

Dim reference26 As Reference
Set reference26 = part1.CreateReferenceFromObject(point2D1)

Dim constraint9 As Constraint
Set constraint9 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference24, reference25, reference26)

constraint9.Mode = catCstModeDrivingDimension

Dim reference27 As Reference
Set reference27 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.1;1);None:();Cf11:());Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;8)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements6 As GeometricElements
Set geometricElements6 = factory2D1.CreateProjections(reference27)

Dim geometry2D5 As Geometry2D
Set geometry2D5 = geometricElements6.Item("Empreinte.1")

geometry2D5.Construction = True

Dim reference28 As Reference
Set reference28 = part1.CreateReferenceFromObject(geometry2D5)

Dim reference29 As Reference
Set reference29 = part1.CreateReferenceFromObject(line2D6)

Dim constraint10 As Constraint
Set constraint10 = constraints1.AddBiEltCst(catCstTypeDistance, reference28, reference29)

constraint10.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint10.Dimension

length1.Value = 2.4

sketch1.CloseEdition

Dim pad2 As Pad
Set pad2 = shapes1.Item("Extrusion.2")

part1.InWorkObject = pad2

part1.Update

length1.Value = 2.4

End Sub


Sub ExtrudeConnectorEnd()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Esquisse.3")

Dim pad1 As Pad
Set pad1 = shapeFactory1.AddNewPad(sketch1, 50#)

Dim limit1 As Limit
Set limit1 = pad1.FirstLimit

Dim length1 As Length
Set length1 = limit1.Dimension

length1.Value = 5#

part1.Update

End Sub

Sub SkecthBaseHole()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(Pad.3;2);None:();Cf11:());Pad.3_ResultOUT;Z0;G4711)")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = -5#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = -1#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("Repère")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("Axe horizontal")

line2D1.ReportName = 5

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("Axe vertical")

line2D2.ReportName = 6

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(-25#, 25#)

point2D1.ReportName = 7

Dim point2D2 As Point2D
Set point2D2 = factory2D1.CreatePoint(-7.5, 42.5)

point2D2.ReportName = 8

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(-7.5, 7.5)

point2D3.ReportName = 9

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(-7.5, 42.5, -7.5, 7.5)

line2D3.ReportName = 1

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D3

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(-42.5, 7.5)

point2D4.ReportName = 10

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(-7.5, 7.5, -42.5, 7.5)

line2D4.ReportName = 2

line2D4.StartPoint = point2D3

line2D4.EndPoint = point2D4

Dim point2D5 As Point2D
Set point2D5 = factory2D1.CreatePoint(-42.5, 42.5)

point2D5.ReportName = 11

Dim line2D5 As Line2D
Set line2D5 = factory2D1.CreateLine(-42.5, 7.5, -42.5, 42.5)

line2D5.ReportName = 3

line2D5.StartPoint = point2D4

line2D5.EndPoint = point2D5

Dim line2D6 As Line2D
Set line2D6 = factory2D1.CreateLine(-42.5, 42.5, -7.5, 42.5)

line2D6.ReportName = 4

line2D6.StartPoint = point2D5

line2D6.EndPoint = point2D2

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(line2D3)

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(line2D2)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeVerticality, reference2, reference3)

constraint1.Mode = catCstModeDrivingDimension

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(line2D4)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference4, reference5)

constraint2.Mode = catCstModeDrivingDimension

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromObject(line2D5)

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(line2D2)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddBiEltCst(catCstTypeVerticality, reference6, reference7)

constraint3.Mode = catCstModeDrivingDimension

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromObject(line2D6)

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(line2D1)

Dim constraint4 As Constraint
Set constraint4 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference8, reference9)

constraint4.Mode = catCstModeDrivingDimension

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromObject(line2D3)

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromObject(line2D5)

Dim reference12 As Reference
Set reference12 = part1.CreateReferenceFromObject(point2D1)

Dim constraint5 As Constraint
Set constraint5 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference10, reference11, reference12)

constraint5.Mode = catCstModeDrivingDimension

Dim reference13 As Reference
Set reference13 = part1.CreateReferenceFromObject(line2D4)

Dim reference14 As Reference
Set reference14 = part1.CreateReferenceFromObject(line2D6)

Dim reference15 As Reference
Set reference15 = part1.CreateReferenceFromObject(point2D1)

Dim constraint6 As Constraint
Set constraint6 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference13, reference14, reference15)

constraint6.Mode = catCstModeDrivingDimension

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim pad1 As Pad
Set pad1 = shapes1.Item("Extrusion.3")

Dim reference16 As Reference
Set reference16 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.3;0:(Brp:(Sketch.3;8)));None:();Cf11:());Face:(Brp:(Pad.3;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateProjections(reference16)

Dim geometry2D1 As Geometry2D
Set geometry2D1 = geometricElements2.Item("Empreinte.1")

geometry2D1.Construction = True

Dim reference17 As Reference
Set reference17 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.3;0:(Brp:(Sketch.3;11)));None:();Cf11:());Face:(Brp:(Pad.3;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements3 As GeometricElements
Set geometricElements3 = factory2D1.CreateProjections(reference17)

Dim geometry2D2 As Geometry2D
Set geometry2D2 = geometricElements3.Item("Empreinte.1")

geometry2D2.Construction = True

Dim reference18 As Reference
Set reference18 = part1.CreateReferenceFromObject(geometry2D1)

Dim reference19 As Reference
Set reference19 = part1.CreateReferenceFromObject(geometry2D2)

Dim reference20 As Reference
Set reference20 = part1.CreateReferenceFromObject(point2D1)

Dim constraint7 As Constraint
Set constraint7 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference18, reference19, reference20)

constraint7.Mode = catCstModeDrivingDimension

Dim reference21 As Reference
Set reference21 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.3;0:(Brp:(Sketch.3;10)));None:();Cf11:());Face:(Brp:(Pad.3;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements4 As GeometricElements
Set geometricElements4 = factory2D1.CreateProjections(reference21)

Dim geometry2D3 As Geometry2D
Set geometry2D3 = geometricElements4.Item("Empreinte.1")

geometry2D3.Construction = True

Dim reference22 As Reference
Set reference22 = part1.CreateReferenceFromBRepName("FEdge:(Edge:(Face:(Brp:(Pad.3;0:(Brp:(Sketch.3;6)));None:();Cf11:());Face:(Brp:(Pad.3;2);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", pad1)

Dim geometricElements5 As GeometricElements
Set geometricElements5 = factory2D1.CreateProjections(reference22)

Dim geometry2D4 As Geometry2D
Set geometry2D4 = geometricElements5.Item("Empreinte.1")

geometry2D4.Construction = True

Dim reference23 As Reference
Set reference23 = part1.CreateReferenceFromObject(geometry2D3)

Dim reference24 As Reference
Set reference24 = part1.CreateReferenceFromObject(geometry2D4)

Dim reference25 As Reference
Set reference25 = part1.CreateReferenceFromObject(point2D1)

Dim constraint8 As Constraint
Set constraint8 = constraints1.AddTriEltCst(catCstTypeEquidistance, reference23, reference24, reference25)

constraint8.Mode = catCstModeDrivingDimension

Dim reference26 As Reference
Set reference26 = part1.CreateReferenceFromObject(line2D6)

Dim constraint9 As Constraint
Set constraint9 = constraints1.AddMonoEltCst(catCstTypeLength, reference26)

constraint9.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint9.Dimension

length1.Value = 35#

Dim reference27 As Reference
Set reference27 = part1.CreateReferenceFromObject(line2D3)

Dim constraint10 As Constraint
Set constraint10 = constraints1.AddMonoEltCst(catCstTypeLength, reference27)

constraint10.Mode = catCstModeDrivingDimension

Dim length2 As Length
Set length2 = constraint10.Dimension

length2.Value = 35#

sketch1.CloseEdition

part1.InWorkObject = pad1

part1.Update

End Sub

Sub ExtrudeBaseHole()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Esquisse.4")

Dim pocket1 As Pocket
Set pocket1 = shapeFactory1.AddNewPocket(sketch1, 5#)

part1.UpdateObject pocket1

Dim limit1 As Limit
Set limit1 = pocket1.FirstLimit

Dim length1 As Length
Set length1 = limit1.Dimension

length1.Value = 76#

part1.Update

End Sub

Sub SketchHoleBaseToPlugFromSketchPlug()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim selection1 As Selection
Set selection1 = partDocument1.Selection

selection1.Clear

Dim part1 As part
Set part1 = partDocument1.part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Esquisse.2")

selection1.Add sketch1

selection1.Copy

Set partDocument1 = CATIA.ActiveDocument

Dim selection2 As Selection
Set selection2 = partDocument1.Selection

selection2.Clear

selection2.Add sketch1

selection2.Paste

Dim sketch2 As Sketch
Set sketch2 = sketches1.Item("Esquisse.5")

part1.InWorkObject = sketch2

Dim factory2D1 As Factory2D
Set factory2D1 = sketch2.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch2.GeometricElements

Dim line2D1 As Line2D
Set line2D1 = geometricElements1.Item("Droite.1")

line2D1.Construction = True

Dim line2D2 As Line2D
Set line2D2 = geometricElements1.Item("Droite.2")

line2D2.Construction = True

Dim line2D3 As Line2D
Set line2D3 = geometricElements1.Item("Droite.3")

line2D3.Construction = True

Dim line2D4 As Line2D
Set line2D4 = geometricElements1.Item("Droite.4")

line2D4.Construction = True

sketch2.CloseEdition

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim pocket1 As Pocket
Set pocket1 = shapes1.Item("Poche.1")

part1.InWorkObject = pocket1

part1.Update

End Sub


Sub ExtrudeHoleBaseToPlug()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Esquisse.5")

Dim pocket1 As Pocket
Set pocket1 = shapeFactory1.AddNewPocket(sketch1, 20#)

Dim limit1 As Limit
Set limit1 = pocket1.FirstLimit

limit1.LimitMode = catUpThruNextLimit

part1.Update

End Sub

Sub DuplicateInYDir()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromName("")

Dim rectPattern1 As RectPattern
Set rectPattern1 = shapeFactory1.AddNewRectPattern(Nothing, 2, 1, 20#, 20#, 1, 1, reference1, reference2, True, True, 0#)

rectPattern1.FirstRectangularPatternParameters = catInstancesandSpacing

rectPattern1.SecondRectangularPatternParameters = catInstancesandSpacing

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim pad1 As Pad
Set pad1 = shapes1.Item("Extrusion.1")

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(Pad.1;0:(Brp:(Sketch.1;4)));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MFBRepVersion_CXR15)", pad1)

rectPattern1.SetFirstDirection reference3

rectPattern1.SetSecondDirection reference3

Dim linearRepartition1 As LinearRepartition
Set linearRepartition1 = rectPattern1.FirstDirectionRepartition

Dim intParam1 As IntParam
Set intParam1 = linearRepartition1.InstancesCount

intParam1.Value = 5

Dim linearRepartition2 As LinearRepartition
Set linearRepartition2 = rectPattern1.FirstDirectionRepartition

Dim length1 As Length
Set length1 = linearRepartition2.Spacing

length1.Value = 50#

intParam1.Value = 5

part1.Update

End Sub

Sub DuplicateInZDir()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromName("")

Dim rectPattern1 As RectPattern
Set rectPattern1 = shapeFactory1.AddNewRectPattern(Nothing, 2, 1, 20#, 20#, 1, 1, reference1, reference2, True, True, 0#)

rectPattern1.FirstRectangularPatternParameters = catInstancesandSpacing

rectPattern1.SecondRectangularPatternParameters = catInstancesandSpacing

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim rectPattern2 As RectPattern
Set rectPattern2 = shapes1.Item("Répétition rectangulaire.1")

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(RectPattern.1_ResultOUT;4-0:(Brp:(Pad.1;0:(Brp:(Sketch.1;4)))));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MFBRepVersion_CXR15)", rectPattern2)

rectPattern1.SetFirstDirection reference3

rectPattern1.SetSecondDirection reference3

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(RectPattern.1_ResultOUT;4-0:(Brp:((Brp:(Pad.3;0:(Brp:(Sketch.3;10)));Brp:(Pad.1;0:(Brp:(Sketch.1;6)))))));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MFBRepVersion_CXR15)", rectPattern2)

rectPattern1.SetFirstDirection reference4

rectPattern1.SetSecondDirection reference4

Dim linearRepartition1 As LinearRepartition
Set linearRepartition1 = rectPattern1.FirstDirectionRepartition

Dim intParam1 As IntParam
Set intParam1 = linearRepartition1.InstancesCount

intParam1.Value = 2

Dim linearRepartition2 As LinearRepartition
Set linearRepartition2 = rectPattern1.FirstDirectionRepartition

Dim length1 As Length
Set length1 = linearRepartition2.Spacing

length1.Value = 50#

intParam1.Value = 2

part1.Update

End Sub


Sub ExtudeConnectorBaseYM()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("")

Dim pad1 As Pad
Set pad1 = shapeFactory1.AddNewPadFromRef(reference1, 20#)

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim rectPattern1 As RectPattern
Set rectPattern1 = shapes1.Item("Répétition rectangulaire.2")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(RectPattern.2_ResultOUT;1-0:(Brp:((Brp:(Pad.3;0:(Brp:(Sketch.3;6)));Brp:(Pad.1;0:(Brp:(Sketch.1;9)))))));None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", rectPattern1)

pad1.SetProfileElement reference2

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(RectPattern.2_ResultOUT;1-0:(Brp:((Brp:(Pad.3;0:(Brp:(Sketch.3;6)));Brp:(Pad.1;0:(Brp:(Sketch.1;9)))))));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", rectPattern1)

pad1.SetDirection reference3

part1.Update

End Sub


Sub ExtudeConnectorBaseYP()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As part
Set part1 = partDocument1.part

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("")

Dim pad1 As Pad
Set pad1 = shapeFactory1.AddNewPadFromRef(reference1, 20#)

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("Corps principal")

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim rectPattern1 As RectPattern
Set rectPattern1 = shapes1.Item("Répétition rectangulaire.2")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(RectPattern.2_ResultOUT;1-0:(Brp:(RectPattern.1_ResultOUT;4-0:(Brp:((Brp:(Pad.3;0:(Brp:(Sketch.3;10)));Brp:(Pad.1;0:(Brp:(Sketch.1;6)))))))));None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", rectPattern1)

pad1.SetProfileElement reference2

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(RectPattern.2_ResultOUT;1-0:(Brp:(RectPattern.1_ResultOUT;4-0:(Brp:((Brp:(Pad.3;0:(Brp:(Sketch.3;10)));Brp:(Pad.1;0:(Brp:(Sketch.1;6)))))))));None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", rectPattern1)

pad1.SetDirection reference3

part1.Update

End Sub



Sub ExtudeConnectorBaseY()

ExtudeConnectorBaseYM
ExtudeConnectorBaseYP
End Sub


Private Sub CommandButton1_Click()
CreatePart
SketchConnectorBase
ExtrudeConnectorBase
SketchConnectorPlug
ExtrudeConnectorPlug
SkecthConnectorEnd
ExtrudeConnectorEnd
SkecthBaseHole
ExtrudeBaseHole

SketchHoleBaseToPlugFromSketchPlug
ExtrudeHoleBaseToPlug

DuplicateInYDir
DuplicateInZDir

ExtudeConnectorBaseY

End Sub


