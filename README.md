# Autocad Set Viewport

Last Edited: Aug 26, 2019 11:09 PM

## Scope

The goal of this tool is to set export a sheet from Revit to dwg and set its plan viewports UCS to world coordinates. 

![](images/Untitled-4b00f463-7562-4b60-b02e-a3ecc1f2ac9c.png)

In doing this, the coordinates of a point inside the viewport will match the shared coordinates of the same point in Revit:

![](images/Untitled-04ef1935-f048-4be3-a2dd-004bbe20b13b.png)

When a sheet is exported from Revit, its viewports are placed close to the origin neither in shared or project base coordinates:

![](images/Untitled-5efca99a-6d9a-4f4a-b183-606c23f48b14.png)

## Workflow

1. Export the sheet from Revit without the plan views

![](images/Untitled-5c6c24bb-cf78-46f4-863f-aa81ffc38560.png)

2. Export the viewports plan views in shared coordinates and import them in the sheet dwg file. If they do not overlap we can bind them, otherwise we can place them on different layers and control their visibility from the viewport layer manager. 

![](images/Untitled-5568eb64-0af0-4b3e-8ff8-9652b57dd7ac.png)

![](images/Untitled-b4e3cba7-1444-4fdb-9469-11342887ddbd.png)

Overlapping plans could also be placed at different Z levels and the visibility controlled by a clip boundary but the clip boundary does not work with text:

![](images/Untitled-a89902fb-e57b-4043-811c-5caaeb5d5969.png)

![](images/Untitled-7008e7cf-ae89-421a-a31d-93b8ab76397f.png)

![](images/Untitled-6ff871bd-c93e-45b3-a822-c5ec5a574b96.png)

3. Move the viewport center to match the center of each plan and rotate the view by the Angle to True North so it looks like what's on the Revit sheet.

![](images/Untitled-9a10dbe2-0e71-4167-8bd1-5463311e1a8e.png)

![](images/Untitled-10039daa-38d7-4bd6-b39e-1df276b40263.png)

## Addins

### Revit

The tool comprises two addins: one for Revit and one for Autocad. 
The Revit addin will add a new tab with 4 buttons 

![](images/Untitled-8daa4f11-4ac3-46d0-9ae1-cb5b42c0f23b.png)

*Check Sheets*: checks if there are plan views on a sheet and if they are overlapping or not. It exports a csv file summary.

![](images/Untitled-e0063903-f223-4d64-893e-14a22666cd0e.png)

![](images/Untitled-a2060c48-49c0-4506-a17d-e64389f1a31e.png)

*Check Viewport Size*: checks that the viewports on a sheet are within the titleblock boundaries. It exports a csv file summary.

![](images/Untitled-83299e12-06f9-482d-9a11-5f1a84a1a1fc.png)

![](images/Untitled-0b69202e-f01f-4f5a-b137-380de1341356.png)

*Export Sheets*: exports the selected sheets to dwg (mm, shared coordinates (not meaningful), no xrefs). Before exporting, it hides all the content of Floor, Ceiling, Area and Engineering plans. (BUG: if a keynote is present on a sheet, the content get hidden too). For output see Workflow step 1.

![](images/Untitled-06ec8aa8-85bd-4202-a6b4-4502af95a570.png)

*Export Xrefs + CSV*: same principle as Export Sheets but working on plan views only. For each selected sheet, loop through its viewports and if a viewport is a Floor, Ceiling, Area or Engineering plan it export it (in mm, shared coordinates, no xrefs).

![](images/Untitled-a234ca0b-a6be-460d-a9d3-a17ea626e8d6.png)

It also calculates: 

- the center point on the sheet of each viewport
- the corresponding point in Shared Coordinates
- the Angle to the True North
- the viewport width and height

and stores these information in a csv file along with sheet number, plan view name and sheet group (no plans, plans overlapping or not. No plans and plans not overlapping xrefs can be bound, while plan overlapping xrefs will not be bound).

![](images/Untitled-12f03be5-a7cb-4436-a4a1-509ff01bb639.png)

### Autocad

The Autocad addin has to be loaded ("NETLOAD") and  launched ("Merge" or "MergeAndBind") through the command line.

![](images/Untitled-9b1f31b8-c423-4c6f-af36-e182f25b72f2.png)

Both commands read the csv file and:

- open a dwg sheet file
- import the xref
- locate the corresponding viewport on the paper space and transform its UCS
- hide the xref in all the other viewports on the sheet
- bind or leave it as xref
- purge, audit and save the file

Note: the script could be improved by loading all the xref on a sheet at the same time. At the moment the script load the xref one by one, each time opening and closing the sheet file.

### Problems encountered:

1. Convert the Viewport center coordinates on a sheet to Project Base Point (and then Survey Point) coordinates.

    Solution: Find the Viewport's view center point which is in PBP coordinates.

2. The Viewport center does not always match the View center. Annotations, Matchlines and Grids can affect the extent of the viewport (hence the position of its center point).
Solution: hide all these categories and find the Viewport center that matches the View center. Then find the View center point in PBP coordinates and translate it by the vector from Viewport original center and the center of the Viewport with the categories hidden.
See https://thebuildingcoder.typepad.com/blog/2018/03/boston-forge-accelerator-and-aligning-plan-views.html